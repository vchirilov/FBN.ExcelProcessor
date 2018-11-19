using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "cpg_hierarchy")]
    public class TreeNode
    {
        [Order(1)] public int Id { get; set; }
        [Order(2)] public int ParentId { get; set; }
        [Order(3)] public string Name { get; set; }
        [Order(4)] public int Lft { get; set; }
        [Order(5)] public int Rgt { get; set; }
    }
}
