// <copyright file="VariableExpressionTreeNode.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

using System.Collections.Generic;

namespace SpreadsheetEngine
{
    /// <summary>
    /// Variable Expression Tree Node
    /// </summary>
    internal class VariableExpressionTreeNode : ExpressionTreeNode
    {
        private Dictionary<string, double> validVariables = new Dictionary<string, double>();

        private string name;

        /// <summary>
        /// Gets or sets Variable Name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VariableExpressionTreeNode"/> class.
        /// </summary>
        /// <param name="newName">New variable name</param>
        /// <param name="variables">Reference to dictionary of current variables</param>
        public VariableExpressionTreeNode(string newName, ref Dictionary<string, double> variables)
        {
            this.name = newName;
            this.validVariables = variables;
        }

        /// <summary>
        /// Overriden Evaluate Method that checks the variables dictionary for previous seen variables
        /// Creates new entry if the variable is not found and initializes value to 0
        /// </summary>
        /// <returns>double value of the variable</returns>
        public override double Evaluate()
        {
            if (!this.validVariables.ContainsKey(this.name))
            {
                this.validVariables[this.name] = 0.0;
            }

            return this.validVariables[this.name];
        }
    }
}
