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
    /// Command to Restore Background Color for Cells
    /// </summary>
    public class RestoreBGColorCommand: ICommand
    {
        private SpreadsheetCell spreadsheetCell;
        private uint prevBGColor;
        private uint newBGColor;

        /// <summary>
        /// Binds cell and colors to the instance of the command
        /// </summary>
        /// <param name="aSpreadsheetCell">Cell to act on</param>
        /// <param name="prevColor">Previous BG color</param>
        /// <param name="newColor">New BG color</param>
        public RestoreBGColorCommand(SpreadsheetCell aSpreadsheetCell, uint prevColor, uint newColor)
        {
            this.spreadsheetCell = aSpreadsheetCell;
            this.prevBGColor = prevColor;
            this.newBGColor = newColor;  
        }

        /// <summary>
        /// Executes the restore color command
        /// </summary>
        public void Execute()
        {
            spreadsheetCell.BGColor = prevBGColor;  
        }

        /// <summary>
        /// Unexecutes the command returning the color to its previous value
        /// </summary>
        public void Unexecute()
        {
            spreadsheetCell.BGColor = newBGColor;
        }

        /// <summary>
        /// Gets the Run-time Command Name
        /// </summary>
        /// <returns>string Name</returns>
        public string CommandName()
        {
            return this.GetType().GetTypeInfo().Name;
        }
    }
}
