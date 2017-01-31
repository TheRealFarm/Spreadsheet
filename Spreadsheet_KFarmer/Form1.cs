// Author: Kyle Farmer
// ID: 11342935

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetEngine;
using System.IO;

namespace Spreadsheet_KFarmer
{
    public partial class Form1 : Form
    {
        //private Spreadsheet Spreadsheet_Engine;
        public Spreadsheet mSheet = new Spreadsheet(50, 26);
        
        public Form1()
        {
            InitializeComponent();
        }

        void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // get row and col indices for cell being editted
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            Cell curCell = mSheet.GetCell(row, col);
            dataGridView1.Rows[row].Cells[col].Value = curCell.Text;
        }

        void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            string updateText;

            Cell sheetCell = mSheet.GetCell(row, col);

            try
            {
                // gather the updated text from the cell
                updateText = dataGridView1.Rows[row].Cells[col].Value.ToString();
            }
            catch (NullReferenceException)
            {
                updateText = "";
            }

            IUndoRedoCmd[] undos = new IUndoRedoCmd[1]; 

            undos[0] = new RestoreTextCmd(sheetCell.Text, sheetCell.Name);

            // set the current selected cells text to the updated text
            sheetCell.Text = updateText;
            // since this function is related to cell text, we add a change in 
            // cell text undo to the internal undo stack
            mSheet.AddUndo(new UndoRedoCollection(undos, "change in cell text"));

            // then set the value of the cell to the current cells internal value
            dataGridView1.Rows[row].Cells[col].Value = sheetCell.Value;

            UpdateToolStrip();
        }

        private void SheetChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Value")
            {
                Cell curCell = sender as Cell;
                if (curCell != null)
                {
                    int row = curCell.RowIndex;
                    int col = curCell.ColIndex;
                    dataGridView1.Rows[row].Cells[col].Value = curCell.Value;
                }
            }
            else if (e.PropertyName == "BGColor")
            {
                // get the current cell from the sender
                Cell curCell = sender as Cell;
                
                if (curCell != null)
                {
                    int row = curCell.RowIndex;
                    int col = curCell.ColIndex;
                    int color = curCell.BGColor;
                    // property implementation for cell color
                    dataGridView1.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(color);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.CellBeginEdit += dataGridView1_CellBeginEdit;
            dataGridView1.CellEndEdit += dataGridView1_CellEndEdit;

            mSheet.SSPropertyChanged += SheetChanged;

            // empties grid before set up of the new grid
            dataGridView1.Columns.Clear();

            const string alphabet = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            // creates columms, one for each letter of the alphabet
            for (int i = 1; i < alphabet.Length; i++)
            {
                string header = "";
                this.dataGridView1.Columns.Add(header, Convert.ToString(alphabet[i]));
            }
            // Add set number of rows
            dataGridView1.Rows.Add(50);
            //// then change the header text for those rows
            //for (int i = 0; i < 50; i++)
            //{
            //    dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
            //    // code idea was taken from stackoverflow: http://stackoverflow.com/questions/710064/adding-text-to-datagridview-row-header
            //} 

            int Row_Num = 1;
            foreach (DataGridViewRow DataGridRow in dataGridView1.Rows)
            {
                DataGridRow.HeaderCell.Value = Convert.ToString(Row_Num++);
            }

            // initialize to false so they can't be accessed until a undo/redo is pushed onto their stack
            undoToolStripMenuItem.Enabled = false;
            redoToolStripMenuItem.Enabled = false;
        }

        private void UpdateToolStrip()
        {
            // looking at the cell menu item...
            ToolStripMenuItem option = menuStrip1.Items[1] as ToolStripMenuItem;

            // iterate through each item in the drop down menu
            foreach (ToolStripItem item in option.DropDownItems)
            {
                if (item.Text.Substring(0, 4) == "Undo")
                {
                    // check if an undo is valid and then set the drop down text
                    item.Enabled = mSheet.validUndo;
                    item.Text = "Undo " + mSheet.undoDescription;
                }
                else if (item.Text.Substring(0, 4) == "Redo")
                {
                    // check if a redo is valid and then set the drop down text
                    item.Enabled = mSheet.validRedo;
                    item.Text = "Redo " + mSheet.redoDescription;
                }
            }
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mSheet.Undo(mSheet);
            UpdateToolStrip();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mSheet.Redo(mSheet);
            UpdateToolStrip();
        }

       private void pickCellBackgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
       {
           // create list of undos to hold the actions
           List<IUndoRedoCmd> undos = new List<IUndoRedoCmd>();

           ColorDialog cellcolor = new ColorDialog();
           // check to see if they picked a color
           if (cellcolor.ShowDialog() == DialogResult.OK)
           {
               int colorchoice = cellcolor.Color.ToArgb();

               foreach (DataGridViewCell dataGridCell in dataGridView1.SelectedCells)
               {
                   // get current cell and add a restore background color to our undos list for that cell
                   Cell curCell = mSheet.GetCell(dataGridCell.RowIndex, dataGridCell.ColumnIndex);
                   undos.Add(new RestoreBGColorCmd(curCell.BGColor, curCell.Name));
                   // this fires our SSPropertyChanged and sends the curCell as the sender
                   curCell.BGColor = colorchoice;
               }
               // add the undos to our stack
               mSheet.AddUndo(new UndoRedoCollection(undos, "change in cell background color"));
               UpdateToolStrip();
           }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML files (*.xml)|*.xml";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // create a stream with write permissions
                Stream stream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write);
                mSheet.Save(stream);
                stream.Dispose();
            }

        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XML files (*.xml)|*.xml";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                mSheet.Load(stream);
                stream.Dispose();
            }

            UpdateToolStrip();
        }

    }
}
