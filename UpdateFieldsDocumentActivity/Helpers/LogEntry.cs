using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateFieldsDocumentActivity.Helpers
{
    internal class LogEntry
    {
        public DateTime DateTime { get; set; } = DateTime.Now;
        public string Message { get; set; }
        public string Exception { get; set; }

        public override string ToString()
        {
            return $"{DateTime:yyyy-MM-dd HH:mm:ss} | {Message} | {(Exception != null ? $"ERROR: {Exception}" : "")}";
        }
    }
}
