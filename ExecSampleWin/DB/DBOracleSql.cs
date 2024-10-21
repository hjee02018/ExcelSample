using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecSampleWin.DB
{
    public class DBOracleSql
    {
        public string SELECT_T_DEVICEMAP_CHECKLIST(ConcurrentDictionary<string, string> param)
        {
            StringBuilder sb = new StringBuilder();

            try
            {
                sb.AppendFormat("SELECT *         ");
                sb.AppendFormat("  FROM T_DEVICEMAP_CHECKLIST                   ");
                sb.AppendFormat(" WHERE 1 = 1                           ");
                if (param.ContainsKey("SITE"))
                {
                    if (param.TryGetValue("SITE", out string site))
                        sb.AppendFormat("   AND SITE = '{0}'             \r\n", site);
                }
                sb.AppendFormat(" ORDER BY PLC_IP ASC");

                return sb.ToString();
            }
            catch (Exception ex)
            {
                //LogUtil.Log(LogUtil._ERROR_LEVEL, this.GetType().Name, ex.ToString());
                return "";
            }
            finally
            {
                sb = null;
            }
        }

    }
}
