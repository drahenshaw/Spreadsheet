// <copyright file="ExpressionTreeNode.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    /// <summary>
    /// Abstract Expression Tree Node
    /// </summary>
    internal abstract class ExpressionTreeNode
    {
        /// <summary>
        /// Evaluate Method that must be overriden by all derived classes
        /// Recursively evaluate the Expression Tree
        /// </summary>
        /// <returns>double</returns>
        public abstract double Evaluate();
    }
}
