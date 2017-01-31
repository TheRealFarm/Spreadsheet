// Author: Kyle Farmer
// ID: 11342935

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CptS322;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace SpreadsheetEngine
{
    public class Spreadsheet
    {
        //Spreadsheet Cell change event
        public event PropertyChangedEventHandler SSPropertyChanged;
        private Cell[,] m_CellsArray;
        // dictionary object where the keys are the cells and the values are a list of cells the current is dependent on
        private Dictionary<string, HashSet<string>> dependencies;
        // private int m_rowsCount, m_columnCount; // member variables for the rows and columns count
        public int RowsCount // property for rows count
        {
            get { return m_CellsArray.GetLength(0); }
        }

        public int ColumnCount // property for column count
        {
            get { return m_CellsArray.GetLength(1); }
        }

        // private undo and redo stacks
        private Stack<UndoRedoCollection> Undos = new Stack<UndoRedoCollection>();
        private Stack<UndoRedoCollection> Redos = new Stack<UndoRedoCollection>();

        // create a cell from the abstract base class
        private class m_Cell : Cell
        {
            public m_Cell (int row, int col) : base(row, col) // using the values passed into the base class
            {
                
            }

            // set the protected string "m_value" in the Cell class with the constructor of the cell in Spreadsheet
            // since the Spreadsheet class is the factory for the cells in this program.
            // this also allows the value property in the Cell class to be simply a getter.
            public void setValue(string newValue)
            {
                m_value = newValue;
            }

        }

        public Spreadsheet(int rows, int cols)
        {
            m_CellsArray = new Cell[rows, cols];
            dependencies = new Dictionary<string, HashSet<string>>();
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    // populate cell array with the number of rows and columns hard coded in the
                    // load event
                    Cell curCell = new m_Cell(i, j);
                    curCell.PropertyChanged += CellPropertyChanged;
                    m_CellsArray[i, j] = curCell;
                }
            }
        }

        #region Undo/Redo Functionality
        // check if undo stack is empty
        public bool validUndo
        {
            get { return Undos.Count != 0; }
        }

        // check if redo stack is empty
        public bool validRedo
        {
            get { return Redos.Count != 0; }
        }

        // string property to get the description for the button
        // of the action being performed if undo is available
        public string undoDescription
        {
            get
            {
                if (validUndo)
                    return Undos.Peek().CmdDescription;
                return "";
            }
        }

        // same logic for redo
        public string redoDescription
        {
            get
            {
                if (validRedo)
                    return Redos.Peek().CmdDescription;
                return "";
            }
        }

        // pushes action on the undos stack, clear the redo stack
        public void AddUndo(UndoRedoCollection undo)
        {
            Undos.Push(undo);
            Redos.Clear();
        }

        // For each undo action, push onto the redo stack and pop off the undo stack
        public void Undo(Spreadsheet sheet)
        {
            UndoRedoCollection undo = Undos.Pop();
            Redos.Push(undo.Exec(sheet));
        }
        // Pop off redo, push action to undo
        public void Redo(Spreadsheet sheet)
        {
            UndoRedoCollection redo = Redos.Pop();
            Undos.Push(redo.Exec(sheet));
        }
        // Clear undo/redo stacks
        public void Clear()
        {
            Undos.Clear();
            Redos.Clear();
        }
        #endregion

        // This function will set the expression variable retrieved from the spreadsheet
        // First parameter is the expression were in
        // Secon parameter is the cell of the expression variable we will set
        private void SetExpVariable(ExpTree exp, string ExpVarName)
        {
            Cell expVarCell = GetCell(ExpVarName);
            double valueToBeSet;
            if (string.IsNullOrEmpty(expVarCell.Value))
            {
                exp.SetVar(expVarCell.Name, 0);
            }
            else if (!double.TryParse(expVarCell.Value, out valueToBeSet))
            {
                exp.SetVar(ExpVarName, 0);
            }
            else
            {
                exp.SetVar(ExpVarName, valueToBeSet);
            }
        }

        public void EvalCell(Cell CellArg)
        {
            m_Cell m_c = CellArg as m_Cell;
            if (string.IsNullOrEmpty(m_c.Text))
            {
                m_c.setValue("");
                SSPropertyChanged(CellArg, new PropertyChangedEventArgs("Value"));
            }
            else if (m_c.Text[0] == '=' && m_c.Text.Length > 1)
            {
                bool flag = false;
                string expStr = m_c.Text.Substring(1);
                // set the expression string and create the tree
                ExpTree exp = new ExpTree(expStr);

                // gets expression variables
                string[] expVars = exp.GetVarNames();

                // sets each variable in the expression
                foreach (string varName in expVars)
                {
                    // check if the cell is valid
                    if (GetCell(varName) == null)
                    {
                        // set value to the error and fire property changed event for the value
                        m_c.setValue("!(Bad Reference)");
                        SSPropertyChanged(m_c, new PropertyChangedEventArgs("Value"));
                        flag = true;
                        break;
                    }
                  
                    // check to see if its a self reference
                    if (varName == m_c.Name)
                    {
                        // set value to the error and fire property changed event for the value
                        m_c.setValue("!(Self Reference)");
                        SSPropertyChanged(m_c, new PropertyChangedEventArgs("Value"));
                        flag = true;
                        break;
                    }

                    // valid variable 
                    SetExpVariable(exp, varName);

                    // check for circular reference after we've determined its a valid reference
                    if (CheckCircularRef(varName, m_c.Name))
                    {
                        // set value to the error and fire property changed event for the value
                        m_c.setValue("!(Circular Reference)");
                        SSPropertyChanged(m_c, new PropertyChangedEventArgs("Value"));
                        updateCells(m_c.Name);
                        flag = true;
                        break;
                    }

                }
                if (flag)
                    return;

                // Sets the expression evaluation to the current cells value and fires the property changed event
                m_c.setValue(exp.Eval().ToString());
                SSPropertyChanged(CellArg, new PropertyChangedEventArgs("Value"));
            }
            else // no equals sign present so change the text
            {
                m_c.setValue(m_c.Text);
                SSPropertyChanged(CellArg, new PropertyChangedEventArgs("Value"));
            }
            if (dependencies.ContainsKey(m_c.Name))
            {
                foreach (string name in dependencies[m_c.Name])
                {
                    EvalCell(name);
                }
            }
        }

        // Function that updates all the cells of a circular reference to 0
        // Excel has the feature set up this way, so I decided to mimic this
        // NOTE:
        // It's not extensive. For example, if cell B1 holds A1+A2, and A1 holds a 
        // circular reference, then B1 will not be changed. But any cells that have a 
        // direct reference (i.e. some cell's value =A1) on A1 will be set to 0
        public void updateCells(string originalCell)
        {
            foreach (var cell in dependencies)
            {
                // we have already editted the text of the cell passed in to "!(Circular Reference)"
                // so we continue to the next item in the dictionary of dependencies
                if (originalCell == cell.Key)
                {
                    continue;
                }
                // get the cell based on the key
                m_Cell dependentCell = GetCell(cell.Key) as m_Cell;
                // set value and fire property changed
                dependentCell.setValue("0");
                SSPropertyChanged(dependentCell, new PropertyChangedEventArgs("Value"));
            }

        }

        /* checks for circular reference, if there is, return true, if not return false */
        public bool CheckCircularRef(string referencedCell, string currentCell)
        {
            // if the cell we are referencing is equal to the current cell, we have found a circular reference
            if (referencedCell == currentCell)
            {
                return true;
            }

            // the current cell has no dependencies, so there cannot be a circular reference
            else if (!dependencies.ContainsKey(currentCell))
            {
                return false;
            }

            else
            {
                foreach (string dependent in dependencies[currentCell])
                {
                    // recursively check each dependent in the current cell for circular references
                    // the foreach loop allows us to iterate through the dependencies and pass each
                    // cell into our checks above
                    if (CheckCircularRef(referencedCell, dependent))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        
        // Using string location of the Cell, attempt to discover cell location
        private void EvalCell(string location)
        {
            EvalCell(GetCell(location));
        }

        // Another overload that uses row and col index to determine the cell
        private void EvalCell(int row, int col)
        {
            EvalCell(GetCell(row, col));
        }

        public void CellPropertyChanged(object sender, PropertyChangedEventArgs e)
        {

            if (e.PropertyName == "Text")
            {
                // Create new cell instance and have it as a sender
                m_Cell m_c = sender as m_Cell;

                // free the cell of dependents
                FreeCell(m_c.Name);

                if ((m_c.Text.Length > 1) && (m_c.Text != "") && (m_c.Text[0] == '='))
                {
                    ExpTree exp = new ExpTree(m_c.Text.Substring(1));
                    trackDependents(m_c.Name, exp.GetVarNames());
                }
                EvalCell(sender as Cell);
            }
            else if (e.PropertyName == "BGColor")
            {
                // pass the sender to get the currrent cell in the sheetchanged function in form1
                // fire property changed event for the spreadsheet to color
                SSPropertyChanged(sender, new PropertyChangedEventArgs("BGColor"));
            }
       
        }

        // track the dependencies of a cell and its referenced variables
        private void trackDependents(string cell, string[] vars)
        {
            foreach (string var in vars)
            {
      
                if (!dependencies.ContainsKey(var))
                {
                    // Build dictionary entry for this variable name
                    dependencies[var] = new HashSet<string>();
                }
                // Add this cel name to dependencies for this variable
                dependencies[var].Add(cell);
            }
        }

        // Frees a specified cell to remove its dependencies 
        private void FreeCell(string cellname)
        {
            List<string> keys = new List<string>();
            foreach (string key in dependencies.Keys)
            {
                if (dependencies[key].Contains(cellname))
                {
                    keys.Add(key);
                }
            }
            foreach (string key in keys)
            {
                HashSet<string> set = dependencies[key];

                // modifying a location with a cell, the cell is removed from dependencies
                if (set.Contains(cellname))
                {
                    set.Remove(cellname);
                }
            }
        }

        /* Returns the Cell at a specified row and column.
         * If row/column is out of scope then null is returned */
        public Cell GetCell(int row, int col)
        {
            if (row > RowsCount || col > ColumnCount)
            {
                return null;
            }
            else
            {
                return m_CellsArray[row, col];
            }
        }

        // Overloaded GetCell function using the column letter and row number to return
        // the cell found at the location
        public Cell GetCell(string location)
        {
            char column = location[0];
            if (!Char.IsLetter(column))
            {
                // first character was not a letter (invalid input), so null is returned
                return null;
            }

            int row;
            if (!int.TryParse(location.Substring(1), out row))
            {
                // no numbers found in the string so the input is invalid and null returned
                return null;
            }
            // Cell result
            Cell resCell;
            try
            {
                // row is 0 indexed, convert column to ASCII
                resCell = GetCell(row - 1, (int)column - 'A');
            }
            catch (Exception)
            {
                //error caught, null returned
                return null;
            }
            // location was valid, resCell returned
            return resCell;
        }

        // Overloaded save function where all the saving of cells happens
        public void Save(XmlWriter writer)
        {
            writer.WriteStartElement("Spreadsheet");

            // create a list of cells that have been changed to save
            var savedcells =
                from Cell cell in m_CellsArray
                where !cell.CheckDefault
                select cell;

            foreach (Cell cell in savedcells)
            {
                writer.WriteStartElement("Cell");
                writer.WriteAttributeString("Name", cell.Name);
                writer.WriteElementString("Text", cell.Text);
                writer.WriteElementString("BGColor", cell.BGColor.ToString());
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        //Saves current sheet as an XML file at user-picked destination
        public bool Save(Stream stream)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.NewLineChars = "\r\n";
            settings.NewLineOnAttributes = false;
            settings.Indent = true;

            //Uses XmlWriter to write our settings to the stream
            XmlWriter writer = XmlWriter.Create(stream, settings);
            // check the writer for null
            if (writer == null)
            {
                return false;
            }
            // not null so send the writer to our overloaded save function
            Save(writer);
            writer.Close();
            return true;
        }
 

        public void Load(Stream stream)
        {
            // instantiate xml document
            XDocument doc = null;

            try
            {
                doc = XDocument.Load(stream);
            }
            catch (Exception)
            {
                return;
            }

            if (doc == null)
            {
                return;
            }
            //Terminate current sheet for new loaded sheet
            TerminateSheet();
            XElement spreadsheetElement = doc.Root;

            // Clear undo/ redo stacks.
            Clear();
            
            if ("Spreadsheet" != spreadsheetElement.Name)
            {
                return;
            }

            foreach (XElement child in spreadsheetElement.Elements("Cell"))
            {
                Cell curCell = GetCell(child.Attribute("Name").Value);

                //skips past the empty cells.
                if (curCell == null)
                {
                    continue;
                }

                // Load and set text.
                var textElement = child.Element("Text");
                if (textElement != null)
                {
                    curCell.Text = textElement.Value;
                }

                //Load and set background color.
                var bgElement = child.Element("BGColor");
                if (bgElement != null)
                {
                    curCell.BGColor = int.Parse(bgElement.Value);
                }
            }
        }

        // Clears spreadsheet of current data, making room for loading
        public void TerminateSheet()
        {
            for (int i = 0; i < RowsCount; i++)
            {
                for (int j = 0; j < ColumnCount; j++)
                {
                    if (!m_CellsArray[i, j].CheckDefault)
                    {
                        m_CellsArray[i, j].Terminate();
                    }
                }
            }
        }


        #region Demo
        //public void Demo()
        //{
        //    // set 50 random cell's to "Hello world!"
        //    Random rand = new Random();
        //    for (int i = 0; i < 50; i++)
        //    {
        //        Cell c = GetCell(rand.Next() % RowsCount, rand.Next() % ColumnCount);
        //        c.Text = "Hello World!";
        //    }
            
        //    // iterate through Column B and print the index number (B#)
        //    for (int i = 0; i < RowsCount; i++)
        //    {
        //        Cell c = GetCell(i, 1); // (i, 1): 1 is used for column B
        //        c.Text = "This is cell B" + (c.RowIndex + 1);
        //    }
            
        //    // set the text of of each row in column A to the value of the Column B rows
        //    for (int i = 0; i < RowsCount; i++)
        //    {
        //        Cell c = GetCell(i, 0);
        //        Cell c2 = GetCell(i, 1);
        //        c.Text = c2.Value;
        //    }
        #endregion


    }
}
