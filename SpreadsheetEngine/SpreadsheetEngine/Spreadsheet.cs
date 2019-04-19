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
        private Dictionary<SpreadsheetCell, HashSet<SpreadsheetCell>> dependentCells;

        /// <summary>
        /// Command Stacks for Implementing Undo and Redo Functionality
        /// </summary>
        private Stack<List<SpreadsheetEngine.ICommand>> undoCommands = new Stack<List<SpreadsheetEngine.ICommand>>();
        private Stack<List<SpreadsheetEngine.ICommand>> redoCommands = new Stack<List<SpreadsheetEngine.ICommand>>();

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
            this.dependentCells = new Dictionary<SpreadsheetCell, HashSet<SpreadsheetCell>>();

            // Nested Loop to Create 2D Array of SpreadsheetCells
            for (uint i = 0; i < this.rows; ++i)
            {
                for (uint j = 0; j < this.columns; ++j)
                {
                    // Instatiate Concrete SpreadsheetCells at Position(i, j)
                    SpreadsheetCell cell = new InstanceSpreadsheetCell(i, j);

                    // Subscribe each Cell to PropertyChanges
                    cell.PropertyChanged += this.OnPropertyChanged;

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
        /// <returns>Matching Spreadsheet Cell, Null for invalid reference</returns>
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
                            // Return a valid cell reference
                            return GetCell(cellRow - 1, cellColumn);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            // Return a invalid cell reference
                            return null;
                        }                     
                    }
                }
            }

            // Empty cell name is also invalid
            return null;
        }


        /// <summary>
        /// Called when a Cell's Property Changes to Update Spreadsheet
        /// </summary>
        /// <param name="sender">Cell whose property changed</param>
        /// <param name="e">Type of property change</param>
        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "CellText")
            {
                // Clear Dictionary of Dependencies
                RemoveCellDependency(sender as SpreadsheetCell);

                // Build Expression Tree and Evaluate Cell Value
                EvaluateSpreadsheetCell(sender as SpreadsheetCell);
            }
            else if (e.PropertyName == "BGColor")
            {
                // Tell UI to Update Cell Color
                SpreadsheetCell cellChanged = sender as SpreadsheetCell;
                CellPropertyChanged(cellChanged, new PropertyChangedEventArgs("BGColorChanged"));
            }
        }

        /// <summary>
        /// Evaluates Cell's Value based on Entered Text
        /// </summary>
        /// <param name="cell"></param>
        private void EvaluateSpreadsheetCell(SpreadsheetCell cell)
        {
            if (cell is SpreadsheetCell currentCell && currentCell != null)
            {
                // Empty Cell
                if (string.IsNullOrEmpty(currentCell.CellText))
                {
                    currentCell.CellValue = string.Empty;
                    this.OnPropertyChanged(cell, "CellChanged");
                }

                // Non-Formula
                else if (currentCell.CellText[0] != '=')
                {
                    currentCell.CellValue = currentCell.CellText;
                    this.OnPropertyChanged(cell, "CellChanged");
                }
                
                // Formula
                else
                {
                    // Build the Expression Tree Based on Cell Text
                    SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree(currentCell.CellText.Substring(1));

                    // Get the Variable Names from the Expression Tree
                    string[] variableNames = expressionTree.GetVariableNames();

                    foreach (string variable in variableNames)
                    {
                        SpreadsheetCell variableCell = this.GetCell(variable);

                        // Check for Bad / Self References
                        if (!CheckValidReference(currentCell, variableCell) || CheckSelfReference(currentCell, variableCell))
                        {
                            return;
                        }

                        // Adjust Variable Values
                        if (variableCell.CellValue != string.Empty && !variableCell.CellValue.Contains(" "))
                        {
                            expressionTree.SetVariable(variable, Convert.ToDouble(variableCell.CellValue));
                        }
                        else
                        {
                            // Set Default Variable Value to 0.0
                            expressionTree.SetVariable(variable, 0.0);
                        }
                    }

                    // Mark Cells Dependent on the One Being Changed
                    AddCellDependency(currentCell, variableNames);

                    // Check Dependent Cells for Circular References
                    foreach (string variable in variableNames)
                    {
                        SpreadsheetCell variableCell = this.GetCell(variable);
                        if (CheckCircularReference(variableCell, currentCell))
                        {
                            currentCell.CellValue = "!(circular reference)";
                            this.OnPropertyChanged(currentCell, "CellChanged");
                            return;
                        }
                    }

                    // If no Errors, Evaluate the Formula and Update the Value
                    currentCell.CellValue = expressionTree.Evaluate().ToString();
                    this.OnPropertyChanged(cell, "CellChanged");                   
                }

                // Update the Dependent Cells of Cell Being Changed
                if (dependentCells.ContainsKey(currentCell))
                {
                    foreach (SpreadsheetCell dependentCell in dependentCells[currentCell])
                    {
                        EvaluateSpreadsheetCell(dependentCell);
                    }
                }
            }           
        }

        /// <summary>
        /// Subscribe Dependent Cells as Listeners
        /// </summary>
        /// <param name="listener"></param>
        /// <param name="variables"></param>
        private void SubscribeDependencies(SpreadsheetCell listener, string[] variables)
        {
            foreach (string variable in variables)
            {
                SpreadsheetCell variableCell = GetCell(variable);

                if (variableCell != null)
                {
                    variableCell.DependencyChanged += listener.OnDependencyChanged;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        private bool CheckValidReference(SpreadsheetCell sender, SpreadsheetCell cell)
        {
            if (cell == null)
            {
                sender.CellValue = "!(bad reference)";
                this.OnPropertyChanged(sender, "CellChanged");
                return false;
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        private bool CheckSelfReference(SpreadsheetCell sender, SpreadsheetCell cell)
        {
            if (cell == sender)
            {
                sender.CellValue = "!(self reference)";
                this.OnPropertyChanged(sender, "CellChanged");
                return true;
            }
            return false;
        }

        /// <summary>
        /// Checks for Circular References in Spreadsheet Cells
        /// </summary>
        /// <param name="variableCell">Cell that current cell depends on</param>
        /// <param name="currentCell">The current cell being changed</param>
        /// <returns></returns>
        public bool CheckCircularReference(SpreadsheetCell variableCell, SpreadsheetCell currentCell)
        {
            // Self-Reference is a form of Circular Reference so return
            if (variableCell == currentCell)
            {
                return true;
            }

            // If current cell is not in its dependent dictionary return
            if (!dependentCells.ContainsKey(currentCell))
            {
                return false;
            }

            foreach (SpreadsheetCell dependentCell in dependentCells[currentCell])
            {
                // Recursively Check Other Cell Dependencies
                if (CheckCircularReference(variableCell, dependentCell))
                {
                    // Circular Reference Found
                    return true;
                }
            }

            // Dependencies Have No Circular References
            return false;
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
        /// <param name="listenerCell">New cell to add</param>
        /// <param name="variables">List of existing dependencies</param>
        private void AddCellDependency(SpreadsheetCell listenerCell, string[] variables)
        {
            foreach (string variable in variables)
            {
                SpreadsheetCell variableCell = this.GetCell(variable);

                if (variableCell != null)
                {
                    if (!this.dependentCells.ContainsKey(variableCell))
                    {
                        this.dependentCells[variableCell] = new HashSet<SpreadsheetCell>();
                    }

                    this.dependentCells[variableCell].Add(listenerCell);
                }                
            }
        }

        /// <summary>
        /// Clears the dependency of a cell from any lists
        /// </summary>
        /// <param name="cell">Cell to clear dependency</param>
        private void RemoveCellDependency(SpreadsheetCell cell)
        {
            foreach (HashSet<SpreadsheetCell> dependencySet in this.dependentCells.Values)
            {
                if (dependencySet.Contains(cell))
                {
                    dependencySet.Remove(cell);
                }
            }
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
            this.dependentCells = new Dictionary<SpreadsheetCell, HashSet<SpreadsheetCell>>();

            // Nested Loop to Create 2D Array of SpreadsheetCells
            for (uint i = 0; i < this.rows; ++i)
            {
                for (uint j = 0; j < this.columns; ++j)
                {
                    // Instatiate Concrete SpreadsheetCells at Position(i, j)
                    SpreadsheetCell cell = new InstanceSpreadsheetCell(i, j);

                    // Subscribe each Cell to PropertyChanges
                    cell.PropertyChanged += this.OnPropertyChanged;

                    // Add Each Cell to the Spreadsheet
                    this.spreadsheet[i, j] = cell;
                }
            }
        }

        /// <summary>
        /// Tells UI Level that Spreadsheet Cells Changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="name"></param>
        public void OnPropertyChanged(object sender, string name)
        {
            this.CellPropertyChanged?.Invoke(sender, new PropertyChangedEventArgs(name));
        }
    }
}
