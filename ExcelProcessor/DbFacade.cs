using ExcelProcessor.Config;
using ExcelProcessor.Helpers;
using ExcelProcessor.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelProcessor
{
    public class DbFacade
    {
        public void Insert<T>(List<T> items) where T : new()
        {
            Console.WriteLine($"{typeof(T).Name} data loading into database...");

            const int BATCH = 100;
            var chunks = GetChunks(items, BATCH);

            //Build column list in INSERT statement
            var columns = string.Empty;
            
            foreach (var prop in AttributeHelper.GetSortedProperties<T>())
                columns += $",{prop.Name}";
            columns = columns.TrimStart(',');


            var modelAttr = (ModelAttribute)typeof(T).GetCustomAttribute(typeof(ModelAttribute));
            var connectionString = AppSettings.GetInstance().connectionString;

            ExecuteNonQuery($"TRUNCATE TABLE {modelAttr.Table};", $"Truncate has failed for table {modelAttr.Table}");

            using (var conn = new MySqlConnection(connectionString))
            {                
                conn.Open();

                foreach (var chunk in chunks)
                {
                    List<string> rows = new List<string>();

                    foreach (var item in chunk)
                    {
                        var pairs = DictionaryFromType(item);
                        var parameters = string.Empty;

                        foreach (var pair in pairs)
                        {
                            var val = pair.Value == null ? null : MySqlHelper.EscapeString(pair.Value.ToString());
                            parameters += $",'{val}'";
                        }
                            

                        rows.Add("(" + parameters.TrimStart(',') + ")");
                    }

                    var text = new StringBuilder($"INSERT INTO {modelAttr.Table} ({columns}) VALUES ");
                    text.Append(string.Join(",", rows));
                    text.Append(";");

                    using (MySqlCommand sqlCommand = new MySqlCommand(text.ToString(), conn))
                    {
                        try
                        {
                            sqlCommand.CommandType = CommandType.Text;
                            sqlCommand.ExecuteNonQuery();
                        }
                        catch (Exception exc)
                        {
                            Console.WriteLine($"Insert has failed: {exc.Message}");
                        }
                    }
                }
            }

            Console.WriteLine($"{typeof(T).Name} loaded.");

            if (new T() is CpgProductHierarchy)
            {
                BuildCpgHierarchy(items as List<CpgProductHierarchy>);                
            }
        }

        private List<CpgProductHierarchyTree> BuildCpgHierarchy(List<CpgProductHierarchy> items)
        {
            int hid = 1;
            var treeNodes = new List<CpgProductHierarchyTree>();

            foreach(var item in items)
                treeNodes.Add(new CpgProductHierarchyTree(item));

            //Set Id,ParentId values for level 1 (CategoryGroup)
            var categoryGroups = treeNodes.Select(x => x.CategoryGroup.Name).Distinct();            

            foreach(var categoryGroup in categoryGroups)
            {
                treeNodes.Where(x => x.CategoryGroup.Name.Equals(categoryGroup)).ToList().ForEach(x => x.CategoryGroup.Id = hid);
                hid++;
            }
                        
            //Set Id,ParentId values for level 2 (SubDivision) 
            var subDivisions = treeNodes.Select(x => x.Subdivision.Name).Distinct();

            foreach (var subDivision in subDivisions)
            {
                treeNodes.Where(x => x.Subdivision.Name.Equals(subDivision))
                    .ToList()
                    .ForEach(x => {
                        x.Subdivision.ParentId = x.CategoryGroup.Id;
                        x.Subdivision.Id = hid;
                        });
                hid++;
            }

            //Set Id,ParentId values for level 2 (Category) 
            var categories = treeNodes.Select(x => x.Category.Name).Distinct();

            foreach (var category in categories)
            {
                treeNodes.Where(x => x.Category.Name.Equals(category))
                    .ToList()
                    .ForEach(x => {
                        x.Category.ParentId = x.Subdivision.Id;
                        x.Category.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 3 (Market) 
            var markets = treeNodes.Select(x => x.Market.Name).Distinct();

            foreach (var market in markets)
            {
                treeNodes.Where(x => x.Market.Name.Equals(market))
                    .ToList()
                    .ForEach(x => {
                        x.Market.ParentId = x.Category.Id;
                        x.Market.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 4 (Sector) 
            var sectors = treeNodes.Select(x => x.Sector.Name).Distinct();

            foreach (var sector in sectors)
            {
                treeNodes.Where(x => x.Sector.Name.Equals(sector))
                    .ToList()
                    .ForEach(x => {
                        x.Sector.ParentId = x.Market.Id;
                        x.Sector.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 5 (SubSector) 
            var subSectors = treeNodes.Select(x => x.SubSector.Name).Distinct();

            foreach (var subSector in subSectors)
            {
                treeNodes.Where(x => x.SubSector.Name.Equals(subSector))
                    .ToList()
                    .ForEach(x => {
                        x.SubSector.ParentId = x.Sector.Id;
                        x.SubSector.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 6 (Segment) 
            var segments = treeNodes.Select(x => x.Segment.Name).Distinct();

            foreach (var segment in segments)
            {
                treeNodes.Where(x => x.Segment.Name.Equals(segment))
                    .ToList()
                    .ForEach(x => {
                        x.Segment.ParentId = x.SubSector.Id;
                        x.Segment.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 7 (ProductForm) 
            var forms = treeNodes.Select(x => x.ProductForm.Name).Distinct();

            foreach (var form in forms)
            {
                treeNodes.Where(x => x.ProductForm.Name.Equals(form))
                    .ToList()
                    .ForEach(x => {
                        x.ProductForm.ParentId = x.Segment.Id;
                        x.ProductForm.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 8 (CPG) 
            var cpgs = treeNodes.Select(x => x.CPG.Name).Distinct();

            foreach (var cpg in cpgs)
            {
                treeNodes.Where(x => x.CPG.Name.Equals(cpg))
                    .ToList()
                    .ForEach(x => {
                        x.CPG.ParentId = x.ProductForm.Id;
                        x.CPG.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 9 (BrandForm) 
            var brands = treeNodes.Select(x => x.BrandForm.Name).Distinct();

            foreach (var brand in brands)
            {
                treeNodes.Where(x => x.BrandForm.Name.Equals(brand))
                    .ToList()
                    .ForEach(x => {
                        x.BrandForm.ParentId = x.CPG.Id;
                        x.BrandForm.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 10 (SizePackForm) 
            var sizePacks = treeNodes.Select(x => x.SizePackForm.Name).Distinct();

            foreach (var sizePack in sizePacks)
            {
                treeNodes.Where(x => x.SizePackForm.Name.Equals(sizePack))
                    .ToList()
                    .ForEach(x => {
                        x.SizePackForm.ParentId = x.BrandForm.Id;
                        x.SizePackForm.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 10 (SizePackFormVariant) 
            var variants = treeNodes.Select(x => x.SizePackFormVariant.Name).Distinct();

            foreach (var variant in variants)
            {
                treeNodes.Where(x => x.SizePackFormVariant.Name.Equals(variant))
                    .ToList()
                    .ForEach(x => {
                        x.SizePackFormVariant.ParentId = x.SizePackForm.Id;
                        x.SizePackFormVariant.Id = hid;
                    });
                hid++;
            }

            return treeNodes;
        }

        private void ExecuteNonQuery(string sqlStatement, string message = "SQL execution has failed")
        {
            var connectionString = AppSettings.GetInstance().connectionString;

            using (var conn = new MySqlConnection(connectionString))
            {
                conn.Open();

                using (MySqlCommand sqlCommand = new MySqlCommand(sqlStatement.ToString(), conn))
                {
                    try
                    {
                        sqlCommand.CommandType = CommandType.Text;
                        sqlCommand.ExecuteNonQuery();
                    }
                    catch (Exception exc)
                    {
                        Console.WriteLine($"{message}: {exc.Message}");
                    }
                }
            }

        }        

        private IEnumerable<List<T>> GetChunks<T>(List<T> source, int size = 10)
        {
            for (int i = 0; i < source.Count; i += size)
            {
                yield return source.GetRange(i, Math.Min(size, source.Count - i));
            }
        }

        private Dictionary<string, object> DictionaryFromType(object customType)
        {
            if (customType == null)
                return new Dictionary<string, object>();

            var props = AttributeHelper.GetSortedProperties(customType);

            Dictionary<string, object> dict = new Dictionary<string, object>();

            foreach (PropertyInfo prop in props)
            {
                object value = prop.GetValue(customType, new object[] { });
                dict.Add(prop.Name, value);
            }
            return dict;
        }
    }
}
