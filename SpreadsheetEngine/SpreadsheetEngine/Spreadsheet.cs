// <copyright file="Spreadsheet.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Xml;
using System.IO;
using System.Reflection;
using System.Collections;

namespace CptS321
{
    /// <summary>
    /// Class to Create and Use Spreadsheets
    /// </summary>
    public class Spreadsheet
    {
        /// <summary>
        /// Private Fields
        /// </summary>
        private uint rows;
        private uint columns;
        private SpreadsheetCell[,] spreadsheet;

        /// <summary>
        /// A dictionary for dependent cells as key, and a list of the cells depending on that cell as value
        /// </summary>
        private Dictionary<SpreadsheetCell, List<SpreadsheetCell>> dependentCells;

        /// <summary>
        /// Command Stacks for Implementing Undo and Redo Functionality
        /// </summary>
        private Stack<List<SpreadsheetEngine.ICommand>> undoCommands = new Stack<List<SpreadsheetEngine.ICommand>>();
        private Stack<List<SpreadsheetEngine.ICommand>> redoCommands = new Stack<List<SpreadsheetEngine.ICommand>>();

        private HashSet<SpreadsheetCell> visitedCells = new HashSet<SpreadsheetCell>();
        private HashSet<string> circularVariables = new HashSet<string>();

        /// <summary>
        /// Set the Delegate for Handling Events
        /// </summary>
        public event PropertyChangedEventHandler CellPropertyChanged = delegate { };

        /// <summary>
        /// Private Class to Instantiate Instances of Abstract Class Spreadsheet Cell
        /// </summary>
        private class InstanceSpreadsheetCell : SpreadsheetCell
        {
            public InstanceSpreadsheetCell(uint row, uint column)
                : base(row, column) { }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
        /// Initalizes a new Spreadsheet Object Composed of References to SpreadsheetCells
        /// </summary>
        /// <param name="numRows">Number of Spreadsheet Rows</param>
        /// <param name="numColumns">Number of Spreadsheet Columns</param>
        public Spreadsheet(uint numRows, uint numColumns)
        {
            this.rows = numRows;
            this.columns = numColumns;
            this.spreadsheet = new SpreadsheetCell[this.rows, this.columns];
            this.dependentCells = new Dictionary<SpreadsheetCell, List<SpreadsheetCell>>();

            // Nested Loop to Create 2D Array of SpreadsheetCells
            for (uint i = 0; i < this.rows; ++i)
            {
                for (uint j = 0; j < this.columns; ++j)
                {
                    // Instatiate Concrete SpreadsheetCells at Position(i, j)
                    SpreadsheetCell cell = new InstanceSpreadsheetCell(i, j);

                    // Subscribe each Cell to PropertyChanges
                    cell.PropertyChanged += this.UpdateSpreadsheetCell;

                    // Add Each Cell to the Spreadsheet
                    this.spreadsheet[i, j] = cell;
                }
            }
        }

        /// <summary>
        /// Gets Number of Spreadsheet Rows - ReadOnly
        /// </summary>
        public uint RowCount
        {
            get { return this.rows; }
        }

        /// <summary>
        /// Gets Number of Spreadsheet Columns - ReadOnly
        /// </summary>
        public uint ColumnCount
        {
            get { return this.columns; }
        }

        /// <summary>
        /// Get a Reference to a SpreadsheetCell from the 2D Array at [row,column]
        /// </summary>
        /// <param name="row">Row to GetCell</param>
        /// <param name="column">Column to GetCell</param>
        /// <returns>Cell at [row,column] or null if not Found</returns>
        public SpreadsheetCell GetCell(uint row, uint column)
        {
            // If OutofBounds throw IndexOutofRangeException
            if (row < 0 || row > 100 || column < 0 || column > 100)
            {
                throw new IndexOutOfRangeException();
            }
            else
            {
                return this.spreadsheet[row, column];
            }
        }

        /// <summary>
        /// Overloaded GetCell to find cell based on string name
        /// </summary>
        /// <param name="cellName">String name</param>
        /// <returns>Matching Spreadsheet Cell</returns>
        public SpreadsheetCell GetCell(string cellName)
        {
            if (cellName != string.Empty)
            {
                if (Char.IsLetter(cellName[0]) && Char.IsUpper(cellName[0]))
                {
                    // Subtract 'A' to get Column Index
                    uint cellColumn = (uint)cellName[0] - 'A';

                    if (UInt32.TryParse(cellName.Substring(1), out uint cellRow))
                    {
                        try
                        {
                            return GetCell(cellRow - 1, cellColumn);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            return null;
                        }                     
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Respond to Events where PropertyChanged Occurred
        /// </summary>
        /// <param name="sender">Spreadsheet Cell where Change Occurred</param>
        /// <param name="e">Details of Change</param>
        private void UpdateSpreadsheetCell(object sender, PropertyChangedEventArgs e)
        {
            // Only Update if Cell Property was Changed
            if (e.PropertyName == "CellText" || e.PropertyName == "CellValue")
            {
                // Cast Sender as a Spreadsheet Cell to Change
                SpreadsheetCell cellChanged = sender as SpreadsheetCell;

                //this.UpdateSpreadsheetCell(cellChanged);
                this.OnPropertyChanged("Text");
                
                
            }
            else if (e.PropertyName == "BGColor")
            {
                SpreadsheetCell cellChanged = sender as SpreadsheetCell;
                CellPropertyChanged(cellChanged, new PropertyChangedEventArgs("BGColorChanged"));
            }
        }

        /// <summary>
        /// Overloaded UpdateSpreadsheetCell Method that uses Spreadsheet Engine to Evaluate Cells
        /// </summary>
        /// <param name="cellChanged">Cell to Update</param>
        private void UpdateSpreadsheetCell(SpreadsheetCell cellChanged)
        {
            string cellName = string.Empty;
            cellName += Convert.ToChar('A' + cellChanged.ColumnIndex);
            cellName += (cellChanged.RowIndex + 1).ToString();

            bool error = false;

            // Clear Previous Cell Dependencies
            this.RemoveCellDependency(cellChanged);

            // If User Deleted Cell Contents Update the Value
            if (string.IsNullOrEmpty(cellChanged.CellText))
            {
                cellChanged.CellValue = string.Empty;
            }
            else if (cellChanged.CellText[0] != '=')
            {
                // If we are not evaluating an expression but a single value
                // See if that value is a number
                if (double.TryParse(cellChanged.CellText, out double value))
                {
                    // Create a new expression tree and evaluate the number
                    SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree(cellChanged.CellText);
                    value = expressionTree.Evaluate();
                    //expressionTree.SetVariable(cellChanged.CellText, value);
                    cellChanged.CellValue = value.ToString();
                }
                else
                {
                    // Otherwise update the test
                    cellChanged.CellValue = cellChanged.CellText;
                }
            }
            else
            {
                // Parse out the equal sign to get the epxression
                string expression = cellChanged.CellText.Substring(1);

                // Build an ExpressionTree
                SpreadsheetEngine.ExpressionTree userExpTree = new SpreadsheetEngine.ExpressionTree(expression);

                // Get the existing variable names
                string[] varNames = userExpTree.GetVariableNames();

                // Loop through each variable name
                foreach (string variable in varNames)
                {
                    // Get the cell and add its value to the dictionary
                    SpreadsheetCell variableCell = this.GetCell(variable);

                    if (variableCell == null)
                    {
                        // Check for an Invalid Reference - Out of Bounds or Invalid Variable Format
                        cellChanged.CellValue = "!(bad reference)";
                        this.CellPropertyChanged(cellChanged, new PropertyChangedEventArgs("CellChanged"));
                        error = true;
                    }
                    else if (cellName == variable)
                    {
                        // Check for Self-Reference for Current Cell
                        cellChanged.CellValue = "!(self reference)";
                        this.CellPropertyChanged(cellChanged, new PropertyChangedEventArgs("CellChanged"));
                        error = true;
                    }

                    // If error state, return before evaluating any expressions
                    if (error)
                    {
                        if (this.dependentCells.ContainsKey(cellChanged))
                        {
                            this.UpdateCellDependency(cellChanged);
                        }
                        return;
                    }

                    // Try to parse out the double value
                    if (double.TryParse(variableCell.CellValue, out double value))
                    {
                        // Set the new variable value
                        userExpTree.SetVariable(variable, value);
                    }

                    this.visitedCells.Add(variableCell);
                    this.circularVariables.Add(variable);


                    // Register current cell changed as listener of each variable cell it references
                    variableCell.DependencyChanged += cellChanged.OnDependencyChanged;

                    //this.circularVariables.Add(variable);
                    //error = this.Circular2(cellName, variableCell);
                }

                // Update the cell value to the result of the expression tree   
                cellChanged.CellValue = userExpTree.Evaluate().ToString();     

                if (CircularReferences(cellChanged))
                {
                    cellChanged.CellValue = "!(circular reference)";
                    this.CellPropertyChanged(cellChanged, new PropertyChangedEventArgs("CellChanged"));
                }          
            }

            // Notify Listeners that the cell was changed
            this.CellPropertyChanged(cellChanged, new PropertyChangedEventArgs("CellChanged"));
            this.visitedCells.Clear();
        }

        /// <summary>
        /// Sets a variable name and value for dictionary of ExpressionTree
        /// </summary>
        /// <param name="expressionTree">ExpressionTree being evaluated</param>
        /// <param name="varName">Name of the variable</param>
        private void SetCellVariable(SpreadsheetEngine.ExpressionTree expressionTree, string varName)
        {
            SpreadsheetCell cell = this.GetCell(varName);

            if (double.TryParse(cell.CellValue, out double cellValue))
            {
                expressionTree.SetVariable(varName, cellValue);
            }
            else
            {
                expressionTree.SetVariable(varName, 0.0);
            }
        }

        /// <summary>
        /// Adds a cell to a list of cells that depend on another cell
        /// </summary>
        /// <param name="cell">New cell to add</param>
        /// <param name="independentCells">List of existing dependencies</param>
        private void AddCellDependency(SpreadsheetCell cell, string[] independentCells)
        {
            foreach (string independent in independentCells)
            {
                SpreadsheetCell spreadsheetCell = this.GetCell(independent);

                if (spreadsheetCell != null)
                {
                    this.dependentCells[spreadsheetCell] = new List<SpreadsheetCell>();
                    this.dependentCells[spreadsheetCell].Add(cell);
                }                
            }
        }

        /// <summary>
        /// Clears the dependency of a cell from any lists
        /// </summary>
        /// <param name="cell">Cell to clear dependency</param>
        private void RemoveCellDependency(SpreadsheetCell cell)
        {
            foreach (List<SpreadsheetCell> dependencyList in this.dependentCells.Values)
            {
                if (dependencyList.Contains(cell))
                {
                    dependencyList.Remove(cell);
                }
            }
        }

        /// <summary>
        /// Updates cells that depend on a given cell
        /// </summary>
        /// <param name="cell">Cell that was updated</param>
        private void UpdateCellDependency(SpreadsheetCell cell)
        {
            foreach (SpreadsheetCell dependentCell in this.dependentCells[cell].ToArray())
            {
                this.UpdateSpreadsheetCell(dependentCell);
            }
        }

        private void UpdateDependentCell(object sender, PropertyChangedEventArgs e)
        {
            SpreadsheetCell myTestCell = sender as SpreadsheetCell;
            this.UpdateSpreadsheetCell(myTestCell);
        }

        /// <summary>
        /// Push a list of commands to the top of the stack
        /// </summary>
        /// <param name="undoList"></param>
        public bool AddUndo(List<SpreadsheetEngine.ICommand> undoList)
        {
            this.undoCommands.Push(undoList);
            return true;
        }

        /// <summary>
        /// Undo the most recent list of commands
        /// </summary>
        public bool Undo()
        {
            // If there is a command to undo
            if (this.undoCommands.Count != 0)
            {
                // Pop the list of commands from the top of the stack
                List<SpreadsheetEngine.ICommand> poppedCommands = this.undoCommands.Pop();

                // For each command in the list, execute the restore command to undo
                foreach (SpreadsheetEngine.ICommand command in poppedCommands)
                {
                    command.Execute();
                }

                // Push the undo commands to the redo command stack
                this.redoCommands.Push(poppedCommands);
            }
            // If the stack becomes empty
            if (this.undoCommands.Count == 0)
            {
                // Return true, signifying to disable the undo button
                return true;
            }

            return false;
        }

        /// <summary>
        /// Pops the most recent list of commands from the stack and unexecutes them
        /// </summary>
        /// <returns>T or F</returns>
        public bool Redo()
        {
            if (!RedoStackIsEmpty())
            {
                List<SpreadsheetEngine.ICommand> poppedCommands = redoCommands.Pop();

                foreach (SpreadsheetEngine.ICommand command in poppedCommands)
                {
                    command.Unexecute();
                }

                this.undoCommands.Push(poppedCommands);
            }
            
            if (this.redoCommands.Count == 0)
            {
                return true;
            }
                       
            return false;
        }

        /// <summary>
        /// Checks if the Redo stack has any elements
        /// </summary>
        /// <returns>T or F</returns>
        public bool RedoStackIsEmpty()
        {
            return this.redoCommands.Count == 0;
        }

        /// <summary>
        /// Checks if the Undo stack has any elements
        /// </summary>
        /// <returns>T or F</returns>
        public bool UndoStackIsEmpty()
        {
            return this.undoCommands.Count == 0;
        }

        /// <summary>
        /// Clears the Redo stack
        /// </summary>
        public void ClearRedoStack()
        {
            this.redoCommands.Clear();
        }

        /// <summary>
        /// Sets the text in the undo menu for the most recent command
        /// </summary>
        /// <returns>string name of command type</returns>
        public string UndoCommandName()
        {
            if (this.undoCommands.Any())
            {
                return undoCommands.Peek().ElementAt(0).CommandName();
            }

            //if (this.undoCommands.Any())
            //{
            //    // Peek top of undo stack for most recent undo commands
            //    var commandList = this.undoCommands.Peek();

            //    // Try to cast the ICommand as a RestoreBGColorCommand
            //    SpreadsheetEngine.RestoreBGColorCommand restoreBGColorCommand = commandList[0] as SpreadsheetEngine.RestoreBGColorCommand;
            //    if (restoreBGColorCommand != null)
            //    {
            //        // If the cast was successful, return the command name
            //        return restoreBGColorCommand.GetType().GetTypeInfo().Name;
            //    }

            //    // Try to cast the ICommand as a RestoreTextCommand
            //    SpreadsheetEngine.RestoreTextCommand restoreTextCommand = commandList[0] as SpreadsheetEngine.RestoreTextCommand;
            //    if (restoreTextCommand != null)
            //    {
            //        // If the cast was successful, return the command name
            //        return restoreTextCommand.GetType().GetTypeInfo().Name;
            //    }
            //}

            // Stack was empty - return empty string
            return string.Empty;
        }

        /// <summary>
        /// Sets the text in the redo menu for the most recent command
        /// </summary>
        /// <returns>string name of command type</returns>
        public string RedoCommandName()
        {
            if (this.redoCommands.Any())
            {
                return this.redoCommands.Peek().ElementAt(0).CommandName();
            }

            return string.Empty;
        }

        /// <summary>
        /// Saves the state of the spreadsheet, closes file stream when finished
        /// </summary>
        /// <param name="xmlOutput">File stream for XML output</param>
        public void SaveSpreadsheet(Stream xmlOutput)
        {
            // Create new XML Writer and document
            XmlWriterSettings writerSettings = new XmlWriterSettings()
            {
                Indent = true
            };
            
            XmlWriter xmlWriter = XmlWriter.Create(xmlOutput, writerSettings);

            // Begin Writing
            xmlWriter.WriteStartDocument();

            // Start the Spreadsheet Element
            xmlWriter.WriteStartElement("Spreadsheet");

            // Save each cell in the spreadsheet
            foreach (InstanceSpreadsheetCell cell in this.spreadsheet)
            {
                cell.SaveCell(xmlWriter);
            }

            // Close XML
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndDocument();
            xmlWriter.Close();
        }

        /// <summary>
        /// Loads spreadsheet values from XML document
        /// </summary>
        /// <param name="xmlInput">Input stream to load saved XML state</param>
        public void LoadSpreadsheet(Stream xmlInput)
        {
            // ReInitialize Spreadsheet
            this.ReInitialize();

            // Create new XML reader from the stream
            XmlReader xmlReader = XmlReader.Create(xmlInput);

            string name = string.Empty;
            uint row = 0, col = 0;
            while (xmlReader.Read())
            {
                if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "SpreadsheetCell")
                {
                    // Get the desired attributes from XML document
                    name = xmlReader.GetAttribute("Name");
                    col = name[0] - (uint)'A';
                    row = Convert.ToUInt32(name.Substring(1)) - 1;

                    // Set the cell attributes that correspond to the document
                    this.spreadsheet[row, col].CellText = xmlReader.GetAttribute("Text");
                    this.spreadsheet[row, col].BGColor = Convert.ToUInt32(xmlReader.GetAttribute("BGColor"));                    
                }
            }

            // Close XML
            xmlReader.Close();
        }

        /// <summary>
        /// Clears spreadsheet before loading from XML to avoid merging
        /// </summary>
        public void ReInitialize()
        {
            // Clear Command Stacks
            this.undoCommands.Clear();
            this.redoCommands.Clear();

            // Reset all Cells in Spreadsheet
            this.spreadsheet = new SpreadsheetCell[this.rows, this.columns];
            this.dependentCells = new Dictionary<SpreadsheetCell, List<SpreadsheetCell>>();

            // Nested Loop to Create 2D Array of SpreadsheetCells
            for (uint i = 0; i < this.rows; ++i)
            {
                for (uint j = 0; j < this.columns; ++j)
                {
                    // Instatiate Concrete SpreadsheetCells at Position(i, j)
                    SpreadsheetCell cell = new InstanceSpreadsheetCell(i, j);

                    // Subscribe each Cell to PropertyChanges
                    cell.PropertyChanged += this.UpdateSpreadsheetCell;

                    // Add Each Cell to the Spreadsheet
                    this.spreadsheet[i, j] = cell;
                }
            }
        }


        public bool CircularReferences(SpreadsheetCell cell)
        {
            if (this.visitedCells.Add(cell) == false)
            {
                return true;
            }
            return false;
        }


        public bool Circular2(string cellName, SpreadsheetCell variableCell)
        {
            if (variableCell.CellText.Contains(cellName))
            {
                return true;
            }
            return false;
        }

        public void OnPropertyChanged(string name)
        {
            this.CellPropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
