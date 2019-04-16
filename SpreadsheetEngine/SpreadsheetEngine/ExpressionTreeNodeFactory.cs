// <copyright file="ExpressionTreeNodeFactory.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SpreadsheetEngine
{
    /// <summary>
    /// Factory Pattern to Create Operator Nodes with no hard-coded values
    /// </summary>
    internal class ExpressionTreeNodeFactory
    {
        private static Dictionary<char, Type> validOperators = new Dictionary<char, Type>();

        private delegate void OnOperator(char op, Type type);

        /// <summary>
        /// Checks dictionary for existing key, value pair
        /// </summary>
        /// <param name="key">Operator character</param>
        /// <returns>True or False</returns>
        public bool IsValidOperator(char key)
        {
            if (validOperators.ContainsKey(key))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Gets Associvity for respective operator from dictionary using Reflection
        /// </summary>
        /// <param name="op">Operator Character</param>
        /// <returns>Associativity enum Left or Right</returns>
        public OperatorExpressionTreeNode.Associativity GetAssociativity(char op)
        {
            OperatorExpressionTreeNode.Associativity associativityValue = OperatorExpressionTreeNode.Associativity.COUNT;
            if (validOperators.ContainsKey(op))
            {
                Type type = validOperators[op];
                PropertyInfo propertyInfo = type.GetProperty("Associativity");
                if (propertyInfo != null)
                {
                    object propertyValue = propertyInfo.GetValue(type);
                    if (propertyValue is OperatorExpressionTreeNode.Associativity)
                    {
                        associativityValue = (OperatorExpressionTreeNode.Associativity)propertyValue;
                    }
                }
            }

            return associativityValue;
        }

        /// <summary>
        /// Gets Precedence for respective operatro from dictionary using Reflection
        /// </summary>
        /// <param name="op">Operator Character</param>
        /// <returns>Precedence level as ushort</returns>
        public ushort GetPrecedence(char op)
        {
            ushort precedenceValue = 0;
            if (validOperators.ContainsKey(op))
            {
                Type type = validOperators[op];
                PropertyInfo propertyInfo = type.GetProperty("Precedence");
                if (propertyInfo != null)
                {
                    object propertyValue = propertyInfo.GetValue(type);
                    if (propertyValue is ushort)
                    {
                        precedenceValue = (ushort)propertyValue;
                    }
                }
            }

            return precedenceValue;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpressionTreeNodeFactory"/> class.
        /// Initializes operator dictionary scanning all assemblies for derived operator classes
        /// </summary>
        public ExpressionTreeNodeFactory()
        {
            this.TraverseAvailableOperators((op, type) => validOperators.Add(op, type));
        }

        /// <summary>
        /// Creates Operator Node derived objects based on character passed in
        /// </summary>
        /// <param name="c">Operator Character</param>
        /// <returns>Dervied Operator Node object</returns>
        public OperatorExpressionTreeNode CreateOperatorNode(char c)
        {
            if (validOperators.ContainsKey(c))
            {
                object operatorNodeObject = System.Activator.CreateInstance(validOperators[c]);
                return (OperatorExpressionTreeNode)operatorNodeObject;
            }

            throw new Exception("Unhandled Operator");
        }

        private void TraverseAvailableOperators(OnOperator onOperator)
        {
            // get the type declaration of OperatorNode
            Type operatorNodeType = typeof(OperatorExpressionTreeNode);

            // Iterate over all loaded assemblies:
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                // Get all types that inherit from our OperatorNode class using LINQ
                IEnumerable<Type> operatorTypes =
                assembly.GetTypes().Where(type => type.IsSubclassOf(operatorNodeType));

                // Iterate over those subclasses of OperatorNode
                foreach (var type in operatorTypes)
                {
                    // for each subclass, retrieve the Operator property
                    PropertyInfo operatorField = type.GetProperty("OperatorChar");
                    if (operatorField != null)
                    {
                        // Get the character of the Operator
                        object value = operatorField.GetValue(type);
                        if (value is char)
                        {
                            char operatorSymbol = (char)value;

                            // And infoke the function passed as parameter
                            // with the operator symbol and the operator class
                            onOperator(operatorSymbol, type);
                        }
                    }
                }
            }
        }
    }
}
