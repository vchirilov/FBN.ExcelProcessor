using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelProcessor.Models
{    
    public class CpgProductHierarchyTree
    {
        public CpgProductHierarchyTree(CpgProductHierarchy cpgProductHierarchy)
        {
            CategoryGroup.Name = cpgProductHierarchy.CategoryGroup;
            Subdivision.Name = cpgProductHierarchy.Subdivision;
            Category.Name = cpgProductHierarchy.Category;
            Market.Name = cpgProductHierarchy.Market;
            Sector.Name = cpgProductHierarchy.Sector;
            SubSector.Name = cpgProductHierarchy.SubSector;
            Segment.Name = cpgProductHierarchy.Segment;
            ProductForm.Name = cpgProductHierarchy.ProductForm;
            CPG.Name = cpgProductHierarchy.CPG;
            BrandForm.Name = cpgProductHierarchy.BrandForm;
            SizePackForm.Name = cpgProductHierarchy.SizePackForm;
            SizePackFormVariant.Name = cpgProductHierarchy.SizePackFormVariant;
        }

        public TreeNode CategoryGroup { get; set; } = new TreeNode();
        public TreeNode Subdivision { get; set; } = new TreeNode();
        public TreeNode Category { get; set; } = new TreeNode();
        public TreeNode Market { get; set; } = new TreeNode();
        public TreeNode Sector { get; set; } = new TreeNode();
        public TreeNode SubSector { get; set; } = new TreeNode();
        public TreeNode Segment { get; set; } = new TreeNode();
        public TreeNode ProductForm { get; set; } = new TreeNode();
        public TreeNode CPG { get; set; } = new TreeNode();
        public TreeNode BrandForm { get; set; } = new TreeNode();
        public TreeNode SizePackForm { get; set; } = new TreeNode();
        public TreeNode SizePackFormVariant { get; set; } = new TreeNode();

        public static List<CpgProductHierarchyTree> GetHierarchy(List<CpgProductHierarchy> items)
        {
            int hid = 1;
            var treeNodes = new List<CpgProductHierarchyTree>();

            foreach (var item in items)
                treeNodes.Add(new CpgProductHierarchyTree(item));

            //Set Id,ParentId values for level 1 (CategoryGroup)
            var categoryGroups = treeNodes.Select(x => x.CategoryGroup.Name).Distinct();

            foreach (var categoryGroup in categoryGroups)
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

        public static List<TreeNode> GetTreeNodes(List<CpgProductHierarchy> items)
        {
            List<TreeNode> result = new List<TreeNode>();
            List<CpgProductHierarchyTree> values = GetHierarchy(items);

            foreach(var value in values)
            {
                result.Add(value.CategoryGroup);
                result.Add(value.Subdivision);
                result.Add(value.Category);
                result.Add(value.Market);
                result.Add(value.Sector);
                result.Add(value.SubSector);
                result.Add(value.Segment);
                result.Add(value.ProductForm);
                result.Add(value.CPG);
                result.Add(value.BrandForm);
                result.Add(value.SizePackForm);
                result.Add(value.SizePackFormVariant);
            }

            //return result.OrderBy(x=>x.Id).ThenBy(x=>x.ParentId).ToList();
            return result;
        }
    }
}
