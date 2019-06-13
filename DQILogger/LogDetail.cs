using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DQILogger.Core
{
    public class LogDetail
    {
        public LogDetail()
        {
            Timestamp = DateTime.Now;
            //     AdditionalInfo = new Dictionary<string, object>();
        }
        public DateTime Timestamp { get; private set; }
        public string Message { get; set; }

    
    // WHERE
    public string TaskName { get; set; }
    public string Layer { get; set; }
    public string Location { get; set; }
    public string Hostname { get; set; }

    // WHO
    public string UserId { get; set; }
    public string UserName { get; set; }
    public int CustomerId { get; set; }
    public string CustomerName { get; set; }

    // EVERYTHING ELSE
    public string CorrelationId { get; set; } // exception shielding from server to client
    public long? ElapsedMilliseconds { get; set; }  // only for performance entries
    public Dictionary<string, object> AdditionalInfo { get; set; }  // catch-all for anything else
    public Exception Exception { get; set; }  // the exception for error logging
}

}