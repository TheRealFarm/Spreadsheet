using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public class RestoreBGColorCmd : IUndoRedoCmd
    {
        private int m_color;
        private string m_name;
        public RestoreBGColorCmd(int cellColor, string cellName)
        {
            m_color = cellColor;
            m_name = cellName;
        }

        // store the previous color from the current cell
        // set the new color
        // return the previous color with associated cell
        public IUndoRedoCmd Exec(Spreadsheet ss)
        {
            Cell curCell = ss.GetCell(m_name);
            int prevColor = curCell.BGColor;
            curCell.BGColor = m_color;
            return new RestoreBGColorCmd(prevColor, m_name);
        }
    }
}
