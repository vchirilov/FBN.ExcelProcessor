using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    public class TreeNode
    {
        public int Id { get; set; }
        public int? ParentId { get; set; }
        public string Name { get; set; }
        public int Left { get; set; }
        public int Right { get; set; }
    }
}
