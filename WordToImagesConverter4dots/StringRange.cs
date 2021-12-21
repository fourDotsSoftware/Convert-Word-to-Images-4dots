using System;
using System.Collections.Generic;
using System.Text;

namespace WordToImagesConverter4dots
{
    public class StringRange
    {
        private string Range = "";

        public StringRange(string stringrange)
        {
            Range = stringrange;
        }

        public bool IsInRange(int k)
        {
            if (Range == string.Empty) return true;

            string[] ranges = Range.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            string kk = k.ToString();

            for (int m = 0; m < ranges.Length; m++)
            {
                if (ranges[m] == kk) return true;

                if (ranges[m].IndexOf("-") > 0)
                {
                    string st = ranges[m].Substring(0, ranges[m].IndexOf("-"));
                    int ist = -1;

                    ist = int.Parse(st);

                    string en = ranges[m].Substring(ranges[m].IndexOf("-") + 1);
                    int ien = -1;

                    ien = int.Parse(en);

                    if (k >= ist && k <= ien)
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }

            return false;
        }
    }
}
