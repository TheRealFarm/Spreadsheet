using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public class RestoreTextCmd : IUndoRedoCmd
    {
        private string m_text, m_name;
        public RestoreTextCmd(string cellText, string cellName)
        {
            m_text = cellText;
            m_name = cellName;
        }

        // store previous text from the current cell
        // place current text in the current cell
        // return previous text associated with the current cell name
        public IUndoRedoCmd Exec(Spreadsheet ss)
        {
            Cell curCell = ss.GetCell(m_name);
            string prevTxt = curCell.Text;
            curCell.Text = m_text;
            return new RestoreTextCmd(prevTxt, m_name);
        }
    }
}
