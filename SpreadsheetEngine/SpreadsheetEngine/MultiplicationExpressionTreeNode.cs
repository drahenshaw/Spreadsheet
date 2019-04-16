// <copyright file="MultiplicationExpressionTreeNode.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    /// <summary>
    /// Multiplication Expression Tree Node
    /// </summary>
    internal class MultiplicationExpressionTreeNode : OperatorExpressionTreeNode
    {
        /// <summary>
        /// Gets Static Multiplication Character
        /// </summary>
        public static char OperatorChar => '*';

        /// <summary>
        /// Gets Precedence level for comparing with other operators
        /// </summary>
        public static ushort Precedence => 6;

        /// <summary>
        /// Gets Associativity for Division Operator
        /// </summary>
        public static new Associativity Associativity => Associativity.LEFT;

        /// <summary>
        /// Overriden Evaluate Method for Expression Tree
        /// </summary>
        /// <returns>Multiplication of left and right subtrees</returns>
        public override double Evaluate()
        {
            return this.Left.Evaluate() * this.Right.Evaluate();
        }
    }
}
