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
    }
}
