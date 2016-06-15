using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWvvm.Entities
{
    public class GoogleHistories : Dictionary<string, GoogleHistory>
    {
        private static GoogleHistories m_AllHistories = new GoogleHistories();
        private GoogleHistories()
        { 
        }

        public static GoogleHistories GetAllHistories
        {
            get
            {
                return m_AllHistories;
            }
        }

        public static GoogleHistory GetByRangeName(string rangeName)
        {
            return m_AllHistories.Values.FirstOrDefault(item => string.Compare(item.RangeName, rangeName, true) == 0);
        }
    }
}
