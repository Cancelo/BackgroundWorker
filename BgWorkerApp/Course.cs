using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BgWorkerApp
{
    class Course
    {
        private string name;
        private string part;
        private string dateCount;
        private string total;

        public string Name { get { return name; } set { name = value; } }
        public string Part { get { return part; } set { part = value; } }
        public string DateCount { get { return dateCount; } set { dateCount = value; } }
        public string Total { get { return total; } set { total = value; } }
    }
}
