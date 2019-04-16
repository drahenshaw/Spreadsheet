// <copyright file="OperatorExpressionTreeNode.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

namespace SpreadsheetEngine
{
    /// <summary>
    /// Operator Expression Tree Node
    /// </summary>
    internal class OperatorExpressionTreeNode : ExpressionTreeNode
    {
        /// <summary>
        /// Enum for Left or Right Operator Associativity
        /// </summary>
        public enum Associativity
        {
            /// <summary>
            /// Left Assiociative Operator
            /// </summary>
            LEFT,

            /// <summary>
            /// Right Associative Operator
            /// </summary>
            RIGHT,

            /// <summary>
            /// Number of Enumerated Values
            /// </summary>
            COUNT
        }

        /// <summary>
        /// Gets or sets character that represents the Operator in the node
        /// </summary>
        public char Operator { get; set; }

        /// <summary>
        /// Gets or sets Left subtree Expression Tree Node
        /// </summary>
        public ExpressionTreeNode Left { get; set; }

        /// <summary>
        /// Gets or sets Right subtree Expression Tree Node
        /// </summary>
        public ExpressionTreeNode Right { get; set; }

        /// <summary>
        /// Overriden Evaluate Method
        /// </summary>
        /// <returns>Default value of 0 for non overriden operators</returns>
        public override double Evaluate()
        {
            return 0;
        }
    }
}
