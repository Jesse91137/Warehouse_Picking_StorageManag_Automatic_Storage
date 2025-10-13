using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automatic_Storage
{
    public class _Logger
    {
        public void LogEvent(string enevt,DateTime date,string sqls)
        {
            string strsql = @"INSERT INTO Automatic_Storage_ButtonLogs(EventName, LogTime,EventSQL) 
                                            VALUES(@EventN, @LogTime, @Eventsql)";
            
            SqlParameter[] sqlParameter = new SqlParameter[]
            {
               new SqlParameter("EventN",enevt ),
               new SqlParameter("LogTime",date ),
               new SqlParameter("Eventsql",sqls )
            };
            db.ExecueNonQuery(strsql, CommandType.Text, sqlParameter);

        }
    }
}
