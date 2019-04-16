// <copyright file="AdditionExpressionTreeNode.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    /// <summary>
    /// Addition Operator Expression Tree Node
    /// </summary>
    internal class AdditionExpressionTreeNode : OperatorExpressionTreeNode
    {
        /// <summary>
        /// Gets Static Addition Character
        /// </summary>
        public static char OperatorChar => '+';

        /// <summary>
        /// Gets Precedence level for comparing with other operators
        /// </summary>
        public static ushort Precedence => 7;

        /// <summary>
        /// Gets Associativity for Addition Operator
        /// </summary>
        public static new Associativity Associativity => Associativity.LEFT;

        /// <summary>
        /// Overriden Evaluate Method for Expression Tree
        /// </summary>
        /// <returns>Addition of left and right subtrees</returns>
        public override double Evaluate()
        {
            return this.Left.Evaluate() + this.Right.Evaluate();
        }
    }
}
