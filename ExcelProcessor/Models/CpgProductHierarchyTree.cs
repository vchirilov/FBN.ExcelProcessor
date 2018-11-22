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

        
        public static List<TreeNode> GetTreeNodes(List<CpgProductHierarchy> items)
        {            
            List<TreeNode> result = new List<TreeNode>();            
            TreeNodeComparer comparer = new TreeNodeComparer();

            List<CpgProductHierarchyTree> values = BuildHierarchy(items);

            result.AddRange(values.Select(x => x.CategoryGroup).Distinct(comparer));
            result.AddRange(values.Select(x => x.Subdivision).Distinct(comparer));
            result.AddRange(values.Select(x => x.Category).Distinct(comparer));
            result.AddRange(values.Select(x => x.Market).Distinct(comparer));
            result.AddRange(values.Select(x => x.Sector).Distinct(comparer));
            result.AddRange(values.Select(x => x.SubSector).Distinct(comparer));
            result.AddRange(values.Select(x => x.Segment).Distinct(comparer));
            result.AddRange(values.Select(x => x.ProductForm).Distinct(comparer));
            result.AddRange(values.Select(x => x.CPG).Distinct(comparer));
            result.AddRange(values.Select(x => x.BrandForm).Distinct(comparer));
            result.AddRange(values.Select(x => x.SizePackForm).Distinct(comparer));
            result.AddRange(values.Select(x => x.SizePackFormVariant).Distinct(comparer));
            
            return result;
        }
        private static List<CpgProductHierarchyTree> BuildHierarchy(List<CpgProductHierarchy> items)
        {
            int hid = 1;
            var treeNodes = new List<CpgProductHierarchyTree>();
            StringComparison ignoreCase = StringComparison.CurrentCultureIgnoreCase;

            foreach (var item in items)
                treeNodes.Add(new CpgProductHierarchyTree(item));

            //Set Id,ParentId values for level 1 (CategoryGroup)
            var categoryGroups = treeNodes.Select(x => x.CategoryGroup.Name).Distinct(StringComparer.CurrentCultureIgnoreCase);

            foreach (var categoryGroup in categoryGroups)
            {
                treeNodes.Where(x => x.CategoryGroup.Name.Equals(categoryGroup, ignoreCase))
                    .ToList()
                    .ForEach(x => {
                        x.CategoryGroup.ParentId = -1;
                        x.CategoryGroup.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 2 (SubDivision) 
            var subDivisions = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower()
            }).Distinct();

            foreach (var subDivision in subDivisions)
            {
                treeNodes.Where(x=> 
                x.CategoryGroup.Name.Equals(subDivision.categoryGroup,StringComparison.CurrentCultureIgnoreCase) 
                && x.Subdivision.Name.Equals(subDivision.subDivision, StringComparison.CurrentCultureIgnoreCase)
                )
                    .ToList()
                    .ForEach(x => {
                        x.Subdivision.ParentId = x.CategoryGroup.Id;
                        x.Subdivision.Id = hid;
                    });
                hid++;
            }
            
            //Set Id,ParentId values for level 2 (Category) 
            var categories = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
            }).Distinct();

            foreach (var category in categories)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(category.categoryGroup, StringComparison.CurrentCultureIgnoreCase) 
                && x.Subdivision.Name.Equals(category.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(category.category, StringComparison.CurrentCultureIgnoreCase)
                )
                    .ToList()
                    .ForEach(x => {
                        x.Category.ParentId = x.Subdivision.Id;
                        x.Category.Id = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 3 (Market) 
            var markets = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower()
            }).Distinct();

            foreach (var market in markets)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(market.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(market.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(market.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(market.market, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.Market.ParentId = x.Category.Id;
                    x.Market.Id = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 4 (Sector) 
            var sectors = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower(),
                sector = x.Sector.Name.ToLower()
            }).Distinct();

            foreach (var sector in sectors)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(sector.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(sector.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(sector.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(sector.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Name.Equals(sector.sector, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.Sector.ParentId = x.Market.Id;
                    x.Sector.Id = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 5 (SubSector) 
            var subSectors = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower(),
                sector = x.Sector.Name.ToLower(),
                subSector = x.SubSector.Name.ToLower()
            }).Distinct();

            foreach (var subSector in subSectors)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(subSector.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(subSector.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(subSector.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(subSector.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Name.Equals(subSector.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Name.Equals(subSector.subSector, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.SubSector.ParentId = x.Sector.Id;
                    x.SubSector.Id = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 6 (Segment) 
            var segments = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower(),
                sector = x.Sector.Name.ToLower(),
                subSector = x.SubSector.Name.ToLower(),
                segment = x.Segment.Name.ToLower()
            }).Distinct();

            foreach (var segment in segments)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(segment.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(segment.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(segment.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(segment.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Name.Equals(segment.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Name.Equals(segment.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Name.Equals(segment.segment, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.Segment.ParentId = x.SubSector.Id;
                    x.Segment.Id = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 7 (ProductForm) 
            var forms = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower(),
                sector = x.Sector.Name.ToLower(),
                subSector = x.SubSector.Name.ToLower(),
                segment = x.Segment.Name.ToLower(),
                form = x.ProductForm.Name.ToLower()
            }).Distinct();

            foreach (var form in forms)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(form.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(form.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(form.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(form.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Name.Equals(form.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Name.Equals(form.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Name.Equals(form.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Name.Equals(form.form, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.ProductForm.ParentId = x.Segment.Id;
                    x.ProductForm.Id = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 8 (CPG) 
            var cpgs = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower(),
                sector = x.Sector.Name.ToLower(),
                subSector = x.SubSector.Name.ToLower(),
                segment = x.Segment.Name.ToLower(),
                form = x.ProductForm.Name.ToLower(),
                cpg = x.CPG.Name.ToLower()
            }).Distinct();

            foreach (var cpg in cpgs)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(cpg.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(cpg.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(cpg.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(cpg.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Name.Equals(cpg.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Name.Equals(cpg.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Name.Equals(cpg.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Name.Equals(cpg.form, StringComparison.CurrentCultureIgnoreCase)
                && x.CPG.Name.Equals(cpg.cpg, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.CPG.ParentId = x.ProductForm.Id;
                    x.CPG.Id = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 9 (BrandForm) 
            var brands = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower(),
                sector = x.Sector.Name.ToLower(),
                subSector = x.SubSector.Name.ToLower(),
                segment = x.Segment.Name.ToLower(),
                form = x.ProductForm.Name.ToLower(),
                cpg = x.CPG.Name.ToLower(),
                brand = x.BrandForm.Name.ToLower()
            }).Distinct();

            foreach (var brand in brands)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(brand.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(brand.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(brand.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(brand.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Name.Equals(brand.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Name.Equals(brand.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Name.Equals(brand.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Name.Equals(brand.form, StringComparison.CurrentCultureIgnoreCase)
                && x.CPG.Name.Equals(brand.cpg, StringComparison.CurrentCultureIgnoreCase)
                && x.BrandForm.Name.Equals(brand.brand, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.BrandForm.ParentId = x.CPG.Id;
                    x.BrandForm.Id = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 10 (SizePackForm) 
            var sizePacks = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower(),
                sector = x.Sector.Name.ToLower(),
                subSector = x.SubSector.Name.ToLower(),
                segment = x.Segment.Name.ToLower(),
                form = x.ProductForm.Name.ToLower(),
                cpg = x.CPG.Name.ToLower(),
                brand = x.BrandForm.Name.ToLower(),
                sizePack = x.SizePackForm.Name.ToLower()
            }).Distinct();

            foreach (var sizePack in sizePacks)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(sizePack.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(sizePack.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(sizePack.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(sizePack.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Name.Equals(sizePack.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Name.Equals(sizePack.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Name.Equals(sizePack.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Name.Equals(sizePack.form, StringComparison.CurrentCultureIgnoreCase)
                && x.CPG.Name.Equals(sizePack.cpg, StringComparison.CurrentCultureIgnoreCase)
                && x.BrandForm.Name.Equals(sizePack.brand, StringComparison.CurrentCultureIgnoreCase)
                && x.SizePackForm.Name.Equals(sizePack.sizePack, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.SizePackForm.ParentId = x.BrandForm.Id;
                    x.SizePackForm.Id = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 10 (SizePackFormVariant) 
            var variants = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Name.ToLower(),
                subDivision = x.Subdivision.Name.ToLower(),
                category = x.Category.Name.ToLower(),
                market = x.Market.Name.ToLower(),
                sector = x.Sector.Name.ToLower(),
                subSector = x.SubSector.Name.ToLower(),
                segment = x.Segment.Name.ToLower(),
                form = x.ProductForm.Name.ToLower(),
                cpg = x.CPG.Name.ToLower(),
                brand = x.BrandForm.Name.ToLower(),
                sizePack = x.SizePackForm.Name.ToLower(),
                variant = x.SizePackFormVariant.Name.ToLower()
            }).Distinct();

            foreach (var variant in variants)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Name.Equals(variant.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Name.Equals(variant.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Name.Equals(variant.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Name.Equals(variant.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Name.Equals(variant.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Name.Equals(variant.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Name.Equals(variant.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Name.Equals(variant.form, StringComparison.CurrentCultureIgnoreCase)
                && x.CPG.Name.Equals(variant.cpg, StringComparison.CurrentCultureIgnoreCase)
                && x.BrandForm.Name.Equals(variant.brand, StringComparison.CurrentCultureIgnoreCase)
                && x.SizePackForm.Name.Equals(variant.sizePack, StringComparison.CurrentCultureIgnoreCase)
                && x.SizePackFormVariant.Name.Equals(variant.variant, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.SizePackFormVariant.ParentId = x.SizePackForm.Id;
                    x.SizePackFormVariant.Id = hid;
                });
                hid++;
            }

            return treeNodes;
        }
    }
}
