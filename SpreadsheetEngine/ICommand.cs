using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    /// <summary>
    /// Command Pattern Interface
    /// </summary>
    public interface ICommand
    {
        /// <summary>
        /// Executes the Command
        /// </summary>
        void Execute();

        /// <summary>
        /// Reverses an executed Command
        /// </summary>
        void Unexecute();

        /// <summary>
        /// Retreives the name of the Run-time Command
        /// </summary>
        /// <returns>Name of the command</returns>
        string CommandName();
    }

}
