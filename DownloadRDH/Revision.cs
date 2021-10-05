using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadRDH
{
    public static class Revision
    {
        public static (DateTime revDate, int rev) GetNextRev(DateTime date, int increment = 1)
        {

            var currRevDate = GetCurrRev(date).revDate;

            var nextRevDate = currRevDate.AddDays(7 * increment);
            var nextRevNum = nextRevDate.Day / 7 - (nextRevDate.Day % 7 == 0 ? 1 : 0);

            return (nextRevDate, nextRevNum);
        }

        public static (DateTime revDate, int rev) GetCurrRev(DateTime date)
        {
            var currRevDate = date;

            do
            {
                currRevDate = currRevDate.AddDays(1);
            } while (currRevDate.DayOfWeek != DayOfWeek.Friday);
            var currRevNum = currRevDate.Day / 7 - (currRevDate.Day % 7 == 0 ? 1 : 0);

            return (currRevDate, currRevNum);
        }
    }
}
