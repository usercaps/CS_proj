using System.Collections.Generic;

namespace TitleGen
{
    public class TableRow
    {
        public string testName { get; set; }
        public List<string> values { get; set; }
    }

    public class TableConfig
    {
        public string name { get; set; }
        public string bookmark { get; set; }
        public List<string> columns { get; set; }
        public List<TableRow> rows { get; set; }
    }

    public class TemplateConfig
    {
        public List<TableConfig> tables { get; set; }
    }
}