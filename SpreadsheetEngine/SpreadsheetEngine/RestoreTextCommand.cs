using CptS321;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace SpreadsheetEngine
{
    /// <summary>
    /// Command to restore text values for cells
    /// </summary>
    public class RestoreTextCommand : ICommand
    {
        private SpreadsheetCell spreadsheetCell;
        private string prevText;
        private string newText;

        /// <summary>
        /// Bind command to cell with its old and new values
        /// </summary>
        /// <param name="aSpreadsheetCell">Cell to act on</param>
        /// <param name="oldText">Old text value</param>
        /// <param name="currentText">New text value</param>
        public RestoreTextCommand(SpreadsheetCell aSpreadsheetCell, string oldText, string currentText)
        {
            this.spreadsheetCell = aSpreadsheetCell;
            this.prevText = oldText;
            this.newText = currentText;
        }

        /// <summary>
        /// Restores text to previous value
        /// </summary>
        public void Execute()
        {
            this.spreadsheetCell.CellText = prevText;
        }

        /// <summary>
        /// Restores text to its new value
        /// </summary>
        public void Unexecute()
        {
            this.spreadsheetCell.CellText = newText;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string CommandName()
        {
            return this.GetType().GetTypeInfo().Name;
        }
    }
}
