// <copyright file="Form1.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace Spreadsheet_David_Henshaw
{
    /// <summary>
    /// Form1
    /// </summary>
    public partial class Form1 : Form
    {
        private CptS321.Spreadsheet spreadsheet;

        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// Spreadsheet Application
        /// </summary>
        public Form1()
        {
            this.InitializeComponent();

            // Expand Header Cells to Show Full Number
            this.dataGridView.RowHeadersWidth = 50;

            // Construct a new Spreadsheet Object and Subscribe to Spreadsheet Changes
            this.spreadsheet = new CptS321.Spreadsheet(50, 26);
            this.spreadsheet.CellPropertyChanged += this.UpdateFormUI;

            // this.dataGridView.CellBeginEdit
            this.dataGridView.CellBeginEdit += this.dataGridView_CellBeginEdit;
            this.dataGridView.CellEndEdit += this.dataGridView_CellEndEdit;
        }

        /// <summary>
        /// Fires when Form is Launched
        /// </summary>
        /// <param name="sender">Main</param>
        /// <param name="e">Event arguments</param>
        private void Form1_Load(object sender, EventArgs e)
        {
            // Clear Columns From Designer
            this.dataGridView.Columns.Clear();

            // Start ASCII values at "A"
            int ascii = 65;

            // Generate Columns A - Z
            for (int i = 0; i < 26; ++i)
            {
                char header = (char)(ascii + i);
                string headerString = header.ToString();

                this.dataGridView.Columns.Add(headerString, headerString);
            }

            // Generate Rows 1-50
            this.dataGridView.Rows.Add(50);

            // Update Titles for Row Index Cells
            for (int i = 0; i < 50; ++i)
            {
                var row = this.dataGridView.Rows[i];
                row.HeaderCell.Value = (i + 1).ToString();
            }
        }

        /// <summary>
        /// Fires when Spreadsheet Cells are Changed
        /// </summary>
        /// <param name="sender">Spreadsheet Class</param>
        /// <param name="e">Event arguments</param>
        private void UpdateFormUI(object sender, PropertyChangedEventArgs e)
        {
            // Update the Cell Values at UI Level
            if (e.PropertyName == "CellChanged")
            {
                CptS321.SpreadsheetCell cellChanged = sender as CptS321.SpreadsheetCell;

                this.dataGridView.Rows[(int)cellChanged.RowIndex].Cells[(int)cellChanged.ColumnIndex].Value = cellChanged.CellValue;
            }
            // Update the Cell Colors at UI Level
            else if (e.PropertyName == "BGColorChanged")
            {
                CptS321.SpreadsheetCell cellChanged = sender as CptS321.SpreadsheetCell;

                if (cellChanged != null)
                {
                    uint cellRow = cellChanged.RowIndex;
                    uint cellColumn = cellChanged.ColumnIndex;

                    uint cellColor = cellChanged.BGColor;
                    System.Drawing.Color color = System.Drawing.Color.FromArgb((int)cellColor);

                    dataGridView.Rows[(int)cellRow].Cells[(int)cellColumn].Style.BackColor = color;                    
                }
            }
        }

        /// <summary>
        /// Fires when User Starts Editing Cells in Data Grid
        /// </summary>
        /// <param name="sender">UI Data Grid View</param>
        /// <param name="e">Event arguments</param>
        private void dataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            uint cellRow = (uint)e.RowIndex;
            uint cellColumn = (uint)e.ColumnIndex;

            CptS321.SpreadsheetCell cellToUpdate = this.spreadsheet.GetCell(cellRow, cellColumn);

            if (cellToUpdate != null)
            {
                this.dataGridView.Rows[(int)cellRow].Cells[(int)cellColumn].Value = cellToUpdate.CellText;
            }
        }

        /// <summary>
        /// Fires when Cell Editing Finished
        /// </summary>
        /// <param name="sender">UI Data Grid</param>
        /// <param name="e">Event arguments</param>
        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Get Cell Index
            uint cellRow = (uint)e.RowIndex;
            uint cellColumn = (uint)e.ColumnIndex;

            // Get Cell Instance
            CptS321.SpreadsheetCell cellToUpdate = this.spreadsheet.GetCell(cellRow, cellColumn);
            string oldText = cellToUpdate.CellText;

            // If Valid Cell
            if (cellToUpdate != null)
            {
                try
                {
                    // Update Cell Text at Index with new Value
                    cellToUpdate.CellText = this.dataGridView.Rows[(int)cellRow].Cells[(int)cellColumn].Value.ToString();
                }
                catch (NullReferenceException)
                {
                    // Catch Empty Values and set as Empty
                    cellToUpdate.CellText = string.Empty;
                }

                // Update Grid View
                if (cellToUpdate.CellText != string.Empty)
                {
                    this.dataGridView.Rows[(int)cellRow].Cells[(int)cellColumn].Value.ToString();
                }

                // Create Undo Command for Text Change and add to List
                List<SpreadsheetEngine.ICommand> textUndoCommand = new List<SpreadsheetEngine.ICommand>();
                textUndoCommand.Add(new SpreadsheetEngine.RestoreTextCommand(cellToUpdate, oldText, cellToUpdate.CellText));

                // Add the Undo Command List to the Undo Stack
                this.spreadsheet.AddUndo(textUndoCommand);

                // Enable the Undo Button for the Created Command
                this.undoToolStripMenuItem.Enabled = true;
                
                // Clear any Pending Redo's Since a new Command was Pushed
                this.spreadsheet.ClearRedoStack();
            }

            // Update Menu Text
            UpdateToolStripMenu();
        }

        /// <summary>
        /// Fires when Cell Demo Button Clicked
        /// </summary>
        /// <param name="sender">Demo Button</param>
        /// <param name="e">Event arguments</param>
        private void demoButton_Click(object sender, EventArgs e)
        {
            Random RNGenerator = new Random();

            for (uint i = 0; i < 50; ++i)
            {
                uint randomRow = Convert.ToUInt32(RNGenerator.Next(0, 50));
                uint randomColumn = Convert.ToUInt32(RNGenerator.Next(0, 26));

                CptS321.SpreadsheetCell randomCell = this.spreadsheet.GetCell(randomRow, randomColumn);
                randomCell.CellText = "C++ or C#?";
            }

            for (uint i = 0; i < 50; ++i)
            {
                this.spreadsheet.GetCell(i, 1).CellText = "This is Cell B" + (i + 1);
                this.spreadsheet.GetCell(i, 0).CellText = "=B" + (i + 1);
            }
        }

        /// <summary>
        /// Fires when Background Color Clicked
        /// </summary>
        /// <param name="sender">UI Color Menu</param>
        /// <param name="e">Event arguments</param>
        private void BackgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Open Color Dialog box for user to select a new color
            ColorDialog colorDialog = new ColorDialog();

            // Store a list of commands for undo / redo
            List<SpreadsheetEngine.ICommand> restoreBGColorCommands = new List<SpreadsheetEngine.ICommand>();

            // When the user selects a new color
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                // Iterate through each of the selected cells
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    // Covert to spreadsheetCell location
                    CptS321.SpreadsheetCell spreadsheetCell = spreadsheet.GetCell((uint)cell.RowIndex, (uint)cell.ColumnIndex);
                    uint oldColor = spreadsheetCell.BGColor;
                    uint newColor = (uint)colorDialog.Color.ToArgb();
                    spreadsheetCell.BGColor = newColor;

                    // Create the undo command
                    SpreadsheetEngine.RestoreBGColorCommand restoreColor = new SpreadsheetEngine.RestoreBGColorCommand(spreadsheetCell, oldColor, newColor);
                    restoreBGColorCommands.Add(restoreColor);
                }

                // Add command list to the stack
                spreadsheet.AddUndo(restoreBGColorCommands);
                this.undoToolStripMenuItem.Enabled = true;

                // Check redo stack, clear if not already empty
                if (!this.spreadsheet.RedoStackIsEmpty())
                {
                    this.spreadsheet.ClearRedoStack();
                }
            }

            // Update button labels and enabled state
            UpdateToolStripMenu();
        }

        /// <summary>
        /// Fires when Undo is Clicked
        /// </summary>
        /// <param name="sender">Undo Button</param>
        /// <param name="e">Option arguments</param>
        private void UndoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Execute the Undo Command
            bool stackBecameEmpty = this.spreadsheet.Undo();
            
            // If the stack became emtpy
            if (stackBecameEmpty)
            {
                // Disable Undo Button
                this.undoToolStripMenuItem.Enabled = false;
            }
            else
            {
                // Enable Undo Button
                this.undoToolStripMenuItem.Enabled = true;
            }

            // Also Check Redo Button and Enable
            if (!spreadsheet.RedoStackIsEmpty())
            {
                this.redoToolStripMenuItem.Enabled = true;
            }

            // Update Menu Text
            UpdateToolStripMenu();
        }

        /// <summary>
        /// What to do when Redo Button is Clicked
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Event Arguments</param>
        private void RedoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Redo the last command executed and check status of stack
            bool stackBecameEmtpy = this.spreadsheet.Redo();

            if (stackBecameEmtpy)
            {
                // Disable the Redo button
                this.redoToolStripMenuItem.Enabled = false;
            }
            else
            {
                // Re-Enable the Redo button
                this.redoToolStripMenuItem.Enabled = true;
            }

            // Re-Enable the Undo Button
            if (!spreadsheet.UndoStackIsEmpty())
            {
                this.undoToolStripMenuItem.Enabled = true;
            }

            // Update Reactive Menu Text
            UpdateToolStripMenu();
        }

        /// <summary>
        /// Updates the titles of tool strip menu items to reflect most recent commands
        /// </summary>
        private void UpdateToolStripMenu()
        {
            // Get the Edit Submenu
            ToolStripMenuItem toolStripMenuItem = this.menuStrip1.Items[1] as ToolStripMenuItem;

            // Change Undo or Redo Title to the most Recent Command Executed 
            foreach (ToolStripMenuItem menuItem in toolStripMenuItem.DropDownItems)
            {
                if (menuItem.Text.Substring(0, 4) == "Undo")
                {
                    menuItem.Text = "Undo " + this.spreadsheet.UndoCommandName();
                }
                else if (menuItem.Text.Substring(0, 4) == "Redo")
                {
                    menuItem.Text = "Redo " + this.spreadsheet.RedoCommandName();
                }
            } 
            
            // Disable the Redo button if there are no more commands left to redo
            if (this.spreadsheet.RedoStackIsEmpty())
            {
                this.redoToolStripMenuItem.Enabled = false;
            }
        }

        /// <summary>
        /// Event when Save button is Clicked - Writes Cells to XML
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Event Args</param>
        private void SaveSpreadsheetButton_Click(object sender, EventArgs e)
        {
            // Open Output File and Save Spreadsheet
            FileStream outputFileStream = File.Open(@"D:\Spreadsheet_David_Henshaw\SpreadsheetEngine\SpreadsheetEngine\saveData2.xml", FileMode.Open);
            this.spreadsheet.SaveSpreadsheet(outputFileStream);

            // Close Outfile
            outputFileStream.Close();
        }

        /// <summary>
        /// Event when Load button is Clicked - Loads XML to Cells
        /// </summary>
        /// <param name="sender">Button</param>
        /// <param name="e">Event Args</param>
        private void LoadSpreadsheetButton_Click(object sender, EventArgs e)
        {
            // Open Input File and Load Spreadsheet
            FileStream inputFileStream = File.Open(@"D:\Spreadsheet_David_Henshaw\SpreadsheetEngine\SpreadsheetEngine\saveData2.xml", FileMode.Open);
            this.spreadsheet.LoadSpreadsheet(inputFileStream);

            // Close Input File
            inputFileStream.Close();
        }
    }
}
