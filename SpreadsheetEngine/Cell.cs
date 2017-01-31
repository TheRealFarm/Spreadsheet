// Author: Kyle Farmer
// ID: 11342935

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public abstract class Cell : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        private readonly int m_rowIndex; // member variable for row index
        public int RowIndex // row read only property
        {
            get { return m_rowIndex; }
        }

        private readonly int m_colIndex; // member variable for column index
        public int ColIndex // column read only property
        {
            get { return m_colIndex; }
        }

        // initialize m_text and m_value so we don't throw a null exception in the RestoreTextCmd
        // this way the restored text will be "" string on the first undo
        protected string m_text = ""; // protected member variable for text
        public string Text // text property typed into the cell
        {
            get { return m_text; }
            set
            {
                if (m_text == value) { return; } // no change then return
                // else update the text member variable and notify the property changed event
                m_text = value;
                PropertyChanged(this, new PropertyChangedEventArgs("Text"));
            }
        }

        protected string m_value = ""; // protected member variable
        public string Value
        {
            get { return m_value; }
        }

        private readonly string cellName = "";
        public string Name
        {
            get { return cellName; }
        }

        protected int BackgroundColor = -1;
        public int BGColor
        {
            get { return BackgroundColor; }
            set
            {
                if (BackgroundColor == value) // color has not changed
                    return;
                else // color has changed
                {
                    BackgroundColor = value;
                    // set color value and fire property changed event
                    PropertyChanged(this, new PropertyChangedEventArgs("BGColor"));
                }
            }
        }

        public bool CheckDefault
        {
            get
            {
                if (string.IsNullOrEmpty(Text) && BGColor == -1)
                {
                    return true;
                }

                return false;
            }
        }

        // Row, column, and cell names will be set in the constructor
        public Cell(int row, int col)
        {
            m_rowIndex = row;
            m_colIndex = col;
            cellName += Convert.ToChar('A' + col);
            cellName += (row + 1).ToString();
        }

        //deletes the cell content
        public void Terminate()
        {
            Text = "";
            BGColor = -1;
        }
    }
}
