// <copyright file="SubtractionExpressionTreeNode.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    /// <summary>
    /// Subtraction Expression Tree Node
    /// </summary>
    internal class SubtractionExpressionTreeNode : OperatorExpressionTreeNode
    {
        /// <summary>
        /// Gets Static Subtraction Character
        /// </summary>
        public static char OperatorChar => '-';

        /// <summary>
        /// Gets Precedence level for comparing with other operators
        /// </summary>
        public static ushort Precedence => 7;

        /// <summary>
        /// Gets Associativity for Subtraction Operator
        /// </summary>
        public static new Associativity Associativity => Associativity.LEFT;

        /// <summary>
        /// Overriden Evaluate Method for Expression Tree
        /// </summary>
        /// <returns>Subtraction of left and right subtrees</returns>
        public override double Evaluate()
        {
            return this.Left.Evaluate() - this.Right.Evaluate();
        }
    }
}
