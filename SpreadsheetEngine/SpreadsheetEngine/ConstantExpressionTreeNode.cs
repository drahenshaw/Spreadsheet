// <copyright file="ConstantExpressionTreeNode.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    /// <summary>
    /// Expression Tree Node that holds a constant.
    /// </summary>
    internal class ConstantExpressionTreeNode : ExpressionTreeNode
    {
        /// <summary>
        /// Value of Constant Node as a Double
        /// </summary>
        private double value;

        /// <summary>
        /// Gets or sets value
        /// </summary>
        public double Value { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConstantExpressionTreeNode"/> class.
        /// </summary>
        /// <param name="newValue">New double value</param>
        public ConstantExpressionTreeNode(double newValue)
        {
            this.Value = newValue;
        }

        /// <summary>
        /// Overriden Evaluate Method for Expression Tree
        /// </summary>
        /// <returns>Returns double value</returns>
        public override double Evaluate()
        {
            return this.Value;
        }
    }
}
