using System;
using System.Collections.Generic;

namespace SharePointListitemManager
{
    public class SharepointList
    {
        public IList<Column> Columns { get; set; }
        public IList<Row> Rows { get; set; }

        public class Column
        {
            public string Name { get; set; }
            public Type Type { get; set; }
        }

        public class Row
        {
            public IList<object> Values { get; set; }
        }
    }
}