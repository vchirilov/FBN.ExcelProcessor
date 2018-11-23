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
            CategoryGroup.Value = cpgProductHierarchy.CategoryGroup;
            Subdivision.Value = cpgProductHierarchy.Subdivision;
            Category.Value = cpgProductHierarchy.Category;
            Market.Value = cpgProductHierarchy.Market;
            Sector.Value = cpgProductHierarchy.Sector;
            SubSector.Value = cpgProductHierarchy.SubSector;
            Segment.Value = cpgProductHierarchy.Segment;
            ProductForm.Value = cpgProductHierarchy.ProductForm;
            CPG.Value = cpgProductHierarchy.CPG;
            BrandForm.Value = cpgProductHierarchy.BrandForm;
            SizePackForm.Value = cpgProductHierarchy.SizePackForm;
            SizePackFormVariant.Value = cpgProductHierarchy.SizePackFormVariant;
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
            var categoryGroups = treeNodes.Select(x => x.CategoryGroup.Value).Distinct(StringComparer.CurrentCultureIgnoreCase);

            foreach (var categoryGroup in categoryGroups)
            {
                treeNodes.Where(x => x.CategoryGroup.Value.Equals(categoryGroup, ignoreCase))
                    .ToList()
                    .ForEach(x => {
                        x.CategoryGroup.ParentId = -1;
                        x.CategoryGroup.Hid = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 2 (SubDivision) 
            var subDivisions = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower()
            }).Distinct();

            foreach (var subDivision in subDivisions)
            {
                treeNodes.Where(x=> 
                x.CategoryGroup.Value.Equals(subDivision.categoryGroup,StringComparison.CurrentCultureIgnoreCase) 
                && x.Subdivision.Value.Equals(subDivision.subDivision, StringComparison.CurrentCultureIgnoreCase)
                )
                    .ToList()
                    .ForEach(x => {
                        x.Subdivision.ParentId = x.CategoryGroup.Hid;
                        x.Subdivision.Hid = hid;
                    });
                hid++;
            }
            
            //Set Id,ParentId values for level 2 (Category) 
            var categories = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
            }).Distinct();

            foreach (var category in categories)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(category.categoryGroup, StringComparison.CurrentCultureIgnoreCase) 
                && x.Subdivision.Value.Equals(category.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(category.category, StringComparison.CurrentCultureIgnoreCase)
                )
                    .ToList()
                    .ForEach(x => {
                        x.Category.ParentId = x.Subdivision.Hid;
                        x.Category.Hid = hid;
                    });
                hid++;
            }

            //Set Id,ParentId values for level 3 (Market) 
            var markets = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower()
            }).Distinct();

            foreach (var market in markets)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(market.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(market.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(market.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(market.market, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.Market.ParentId = x.Category.Hid;
                    x.Market.Hid = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 4 (Sector) 
            var sectors = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower(),
                sector = x.Sector.Value.ToLower()
            }).Distinct();

            foreach (var sector in sectors)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(sector.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(sector.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(sector.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(sector.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Value.Equals(sector.sector, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.Sector.ParentId = x.Market.Hid;
                    x.Sector.Hid = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 5 (SubSector) 
            var subSectors = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower(),
                sector = x.Sector.Value.ToLower(),
                subSector = x.SubSector.Value.ToLower()
            }).Distinct();

            foreach (var subSector in subSectors)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(subSector.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(subSector.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(subSector.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(subSector.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Value.Equals(subSector.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Value.Equals(subSector.subSector, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.SubSector.ParentId = x.Sector.Hid;
                    x.SubSector.Hid = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 6 (Segment) 
            var segments = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower(),
                sector = x.Sector.Value.ToLower(),
                subSector = x.SubSector.Value.ToLower(),
                segment = x.Segment.Value.ToLower()
            }).Distinct();

            foreach (var segment in segments)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(segment.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(segment.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(segment.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(segment.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Value.Equals(segment.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Value.Equals(segment.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Value.Equals(segment.segment, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.Segment.ParentId = x.SubSector.Hid;
                    x.Segment.Hid = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 7 (ProductForm) 
            var forms = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower(),
                sector = x.Sector.Value.ToLower(),
                subSector = x.SubSector.Value.ToLower(),
                segment = x.Segment.Value.ToLower(),
                form = x.ProductForm.Value.ToLower()
            }).Distinct();

            foreach (var form in forms)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(form.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(form.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(form.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(form.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Value.Equals(form.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Value.Equals(form.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Value.Equals(form.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Value.Equals(form.form, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.ProductForm.ParentId = x.Segment.Hid;
                    x.ProductForm.Hid = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 8 (CPG) 
            var cpgs = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower(),
                sector = x.Sector.Value.ToLower(),
                subSector = x.SubSector.Value.ToLower(),
                segment = x.Segment.Value.ToLower(),
                form = x.ProductForm.Value.ToLower(),
                cpg = x.CPG.Value.ToLower()
            }).Distinct();

            foreach (var cpg in cpgs)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(cpg.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(cpg.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(cpg.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(cpg.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Value.Equals(cpg.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Value.Equals(cpg.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Value.Equals(cpg.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Value.Equals(cpg.form, StringComparison.CurrentCultureIgnoreCase)
                && x.CPG.Value.Equals(cpg.cpg, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.CPG.ParentId = x.ProductForm.Hid;
                    x.CPG.Hid = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 9 (BrandForm) 
            var brands = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower(),
                sector = x.Sector.Value.ToLower(),
                subSector = x.SubSector.Value.ToLower(),
                segment = x.Segment.Value.ToLower(),
                form = x.ProductForm.Value.ToLower(),
                cpg = x.CPG.Value.ToLower(),
                brand = x.BrandForm.Value.ToLower()
            }).Distinct();

            foreach (var brand in brands)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(brand.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(brand.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(brand.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(brand.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Value.Equals(brand.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Value.Equals(brand.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Value.Equals(brand.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Value.Equals(brand.form, StringComparison.CurrentCultureIgnoreCase)
                && x.CPG.Value.Equals(brand.cpg, StringComparison.CurrentCultureIgnoreCase)
                && x.BrandForm.Value.Equals(brand.brand, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.BrandForm.ParentId = x.CPG.Hid;
                    x.BrandForm.Hid = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 10 (SizePackForm) 
            var sizePacks = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower(),
                sector = x.Sector.Value.ToLower(),
                subSector = x.SubSector.Value.ToLower(),
                segment = x.Segment.Value.ToLower(),
                form = x.ProductForm.Value.ToLower(),
                cpg = x.CPG.Value.ToLower(),
                brand = x.BrandForm.Value.ToLower(),
                sizePack = x.SizePackForm.Value.ToLower()
            }).Distinct();

            foreach (var sizePack in sizePacks)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(sizePack.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(sizePack.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(sizePack.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(sizePack.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Value.Equals(sizePack.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Value.Equals(sizePack.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Value.Equals(sizePack.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Value.Equals(sizePack.form, StringComparison.CurrentCultureIgnoreCase)
                && x.CPG.Value.Equals(sizePack.cpg, StringComparison.CurrentCultureIgnoreCase)
                && x.BrandForm.Value.Equals(sizePack.brand, StringComparison.CurrentCultureIgnoreCase)
                && x.SizePackForm.Value.Equals(sizePack.sizePack, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.SizePackForm.ParentId = x.BrandForm.Hid;
                    x.SizePackForm.Hid = hid;
                });
                hid++;
            }

            //Set Id,ParentId values for level 10 (SizePackFormVariant) 
            var variants = treeNodes.Select(x => new
            {
                categoryGroup = x.CategoryGroup.Value.ToLower(),
                subDivision = x.Subdivision.Value.ToLower(),
                category = x.Category.Value.ToLower(),
                market = x.Market.Value.ToLower(),
                sector = x.Sector.Value.ToLower(),
                subSector = x.SubSector.Value.ToLower(),
                segment = x.Segment.Value.ToLower(),
                form = x.ProductForm.Value.ToLower(),
                cpg = x.CPG.Value.ToLower(),
                brand = x.BrandForm.Value.ToLower(),
                sizePack = x.SizePackForm.Value.ToLower(),
                variant = x.SizePackFormVariant.Value.ToLower()
            }).Distinct();

            foreach (var variant in variants)
            {
                treeNodes.Where(x =>
                x.CategoryGroup.Value.Equals(variant.categoryGroup, StringComparison.CurrentCultureIgnoreCase)
                && x.Subdivision.Value.Equals(variant.subDivision, StringComparison.CurrentCultureIgnoreCase)
                && x.Category.Value.Equals(variant.category, StringComparison.CurrentCultureIgnoreCase)
                && x.Market.Value.Equals(variant.market, StringComparison.CurrentCultureIgnoreCase)
                && x.Sector.Value.Equals(variant.sector, StringComparison.CurrentCultureIgnoreCase)
                && x.SubSector.Value.Equals(variant.subSector, StringComparison.CurrentCultureIgnoreCase)
                && x.Segment.Value.Equals(variant.segment, StringComparison.CurrentCultureIgnoreCase)
                && x.ProductForm.Value.Equals(variant.form, StringComparison.CurrentCultureIgnoreCase)
                && x.CPG.Value.Equals(variant.cpg, StringComparison.CurrentCultureIgnoreCase)
                && x.BrandForm.Value.Equals(variant.brand, StringComparison.CurrentCultureIgnoreCase)
                && x.SizePackForm.Value.Equals(variant.sizePack, StringComparison.CurrentCultureIgnoreCase)
                && x.SizePackFormVariant.Value.Equals(variant.variant, StringComparison.CurrentCultureIgnoreCase)
                )
                .ToList()
                .ForEach(x => {
                    x.SizePackFormVariant.ParentId = x.SizePackForm.Hid;
                    x.SizePackFormVariant.Hid = hid;
                });
                hid++;
            }

            return treeNodes;
        }
    }
}
