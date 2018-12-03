using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "cpg_hierarchy")]
    public class TreeNode
    {
        [Order(1)] public int Hid { get; set; }
        [Order(2)] public int ParentId { get; set; }
        [Order(3)] public string Value { get; set; }
        [Order(4)] public int Lft { get; set; }
        [Order(5)] public int Rgt { get; set; }
    }

    public class TreeNodeComparer : IEqualityComparer<TreeNode>
    {
        public bool Equals(TreeNode x, TreeNode y)
        {
            if (x.Hid == y.Hid && x.ParentId == y.ParentId && x.Value.Equals(y.Value,StringComparison.CurrentCultureIgnoreCase))
                return true;

            return false;
        }
        public int GetHashCode(TreeNode x)
        {
            return 0;
        }

    }
}
