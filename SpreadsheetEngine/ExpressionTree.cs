// <copyright file="ExpressionTree.cs" company="David Henshaw 11215398">
// Copyright (c) David Henshaw 11215398. All rights reserved.
// </copyright>

using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadsheetEngine
{
    /// <summary>
    /// Expression Tree Class that parses a string to evaluate as a double
    /// </summary>
    public class ExpressionTree
    {
        private static ExpressionTreeNodeFactory operatorFactory = new ExpressionTreeNodeFactory();
        private ExpressionTreeNode root;
        private Dictionary<string, double> variables = new Dictionary<string, double>();
        private List<string> expressionParse = new List<string>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpressionTree"/> class.
        /// Parses and Expression string and builds the resulting tree
        /// </summary>
        /// <param name="expression">string expression</param>
        public ExpressionTree(string expression)
        {
            this.variables = new Dictionary<string, double>();
            this.expressionParse = this.ConvertToPostfix(expression);
            this.root = this.ConstructPostfixTree(expressionParse);
        }

        /// <summary>
        /// Changes the variable value associated with its name (key)
        /// </summary>
        /// <param name="variableName">Name of the variable in dictionary</param>
        /// <param name="variableValue">Value that variable should be set to</param>
        public void SetVariable(string variableName, double variableValue)
        {
            this.variables[variableName] = variableValue;
        }

        /// <summary>
        /// Gets Variable names from Variable dictionary in expression tree
        /// </summary>
        /// <returns>String array of keys</returns>
        public string[] GetVariableNames()
        {
            return this.variables.Keys.ToArray();
        }

        /// <summary>
        /// Evaluates the Expression Tree recursively starting with the root
        /// </summary>
        /// <returns>double value as final result of expression tree or NaN</returns>
        public double Evaluate()
        {
            if (this.root != null)
            {
                try
                {
                    return this.root.Evaluate();
                }
                catch (System.DivideByZeroException)
                {
                    System.Console.WriteLine("Expression tree divides by zero, enter a valid expression.");
                    return double.NaN;
                }
            }
            else
            {
                return double.NaN;
            }
        }

        /// <summary>
        /// Builds either a Variable or Constant Expression Tree Node
        /// </summary>
        /// <param name="expression">string to evaluate</param>
        /// <returns>Derived ExpressionTreeNode object</returns>
        private ExpressionTreeNode BuildSimpleNode(string expression)
        {
            double number;

            if (double.TryParse(expression, out number))
            {
                return new ConstantExpressionTreeNode(number);
            }

            this.variables[expression] = 0.0;

            return new VariableExpressionTreeNode(expression, ref this.variables);
        }       

        /// <summary>
        /// Converts infix expression to postfix expression
        /// </summary>
        /// <param name="expression">infix expression</param>
        /// <returns>Postfix expression as list of strings</returns>
        private List<string> ConvertToPostfix(string expression)
        {
            // Dijkstra's Shunting Yard Algorithm

            // Save output as a list of substrings in postfix notation
            List<string> postfixExpression = new List<string>();
            Stack<char> operatorStack = new Stack<char>();

            // Remove whitespace from infix expression
            expression = expression.Replace(" ", string.Empty);

            // Set start point for new operand
            int operandStartIndex = -1;

            // Iterate per character to decide what to do
            for (int i = 0; i < expression.Length; i++)
            {
                // If character belongs to a new operand, set the starting point
                if (!this.IsOperatorOrParenthesis(expression[i]))
                {
                    if (operandStartIndex == -1)
                    {
                        operandStartIndex = i;
                    }
                }
                else
                {
                    // If character is an operator or parenthesis push any started operand to the list
                    if (operandStartIndex != -1)
                    {
                        // Operand has ended and should be added to the list from start point to current position
                        postfixExpression.Add(expression.Substring(operandStartIndex, i - operandStartIndex));

                        // Reset start point for next operand
                        operandStartIndex = -1;
                    }

                    // If character is left parenthesis, push to the stack
                    if (this.IsLeftParenthesis(expression[i]))
                    {
                        operatorStack.Push(expression[i]);
                    }

                    // If character is a right parenthesis
                    else if (this.IsRightParenthesis(expression[i]))
                    {
                        // Discard right parenthesis and pop from stack until left parenthesis is reached
                        while (!this.IsLeftParenthesis(operatorStack.Peek()))
                        {
                            // Add popped operators as substrings to the list
                            postfixExpression.Add(operatorStack.Pop().ToString());
                        }

                        // Discard the matching left parenthesis
                        operatorStack.Pop();
                    }

                    // If character is an operator
                    else if (operatorFactory.IsValidOperator(expression[i]))
                    {
                        // If the stack is empty or top is left parenthesis
                        if (!operatorStack.Any() || this.IsLeftParenthesis(operatorStack.Peek()))
                        {
                            operatorStack.Push(expression[i]);
                        }

                        // If top of stack is an operator
                        else if (operatorFactory.IsValidOperator(operatorStack.Peek()))
                        {
                            ushort expressionPrecedence = operatorFactory.GetPrecedence(expression[i]);
                            ushort stackPrecedence = operatorFactory.GetPrecedence(operatorStack.Peek());
                            OperatorExpressionTreeNode.Associativity expressionAssociativity = operatorFactory.GetAssociativity(expression[i]);

                            // If incoming operator has higher precdence OR equal precedence AND is Right Associative: push the incoming operator to the stack
                            if (expressionPrecedence < stackPrecedence || (expressionPrecedence == stackPrecedence && expressionAssociativity == OperatorExpressionTreeNode.Associativity.RIGHT))
                            {
                                operatorStack.Push(expression[i]);
                            }
                            else
                            {
                                // If the incoming operator has lower precedence OR equal precedence AND is Left Associative pop from the stack until this is no longer true
                                while  (operatorStack.Any() &&
                                      ((operatorFactory.IsValidOperator(operatorStack.Peek()) &&
                                        expressionPrecedence > operatorFactory.GetPrecedence(operatorStack.Peek())) ||
                                       (expressionPrecedence == operatorFactory.GetPrecedence(operatorStack.Peek()) &&
                                        expressionAssociativity == OperatorExpressionTreeNode.Associativity.LEFT)))
                                {
                                    // Add popped operators as substrings to the list
                                    postfixExpression.Add(operatorStack.Pop().ToString());
                                }

                                // Push the incoming operator after the stack has been adjusted
                                operatorStack.Push(expression[i]);
                            }
                        }
                    }
                }
            }

            // If at the end of the expression an operand was started, push it to the output list
            if (operandStartIndex != -1)
            {
                postfixExpression.Add(expression.Substring(operandStartIndex, expression.Length - operandStartIndex));
            }

            // At the end of the expression, pop all remaining symbols from the stack
            while (operatorStack.Count > 0)
            {
                postfixExpression.Add(operatorStack.Pop().ToString());
            }

            // Returns postfix expression as a list of substrings
            return postfixExpression;
        }

        /// <summary>
        /// Constructs a new Expression Tree for a postfix Expression
        /// </summary>
        /// <param name="postfixParse">Postfix expression string</param>
        /// <returns>Root of the constructed tree</returns>
        private ExpressionTreeNode ConstructPostfixTree(List<string> postfixParse)
        {
            Stack<ExpressionTreeNode> treeStack = new Stack<ExpressionTreeNode>();            

            for (int i = 0; i < postfixParse.Count; i++)
            {
                // If operator, pop twice and from stack and create new tree
                if (operatorFactory.IsValidOperator(postfixParse[i].ElementAt(0)))
                {
                    ExpressionTreeNode newRight = treeStack.Pop();
                    ExpressionTreeNode newLeft = treeStack.Pop();
                    OperatorExpressionTreeNode newRoot = operatorFactory.CreateOperatorNode(postfixParse[i].ElementAt(0));

                    newRoot.Right = newRight;
                    newRoot.Left = newLeft;

                    treeStack.Push(newRoot);
                }
                else
                {
                    // Create a new variable or constant node and push to the stack
                    treeStack.Push(this.BuildSimpleNode(postfixParse[i]));
                }
            }

            if (treeStack.Count == 1)
            {
                return treeStack.Pop();
            }
            else
            {
                // Throw exception?
                return null;
            }
        }        

        /// <summary>
        /// Checks if character exists in operator dictionary or is a parenthesis
        /// </summary>
        /// <param name="key">Character to check</param>
        /// <returns>bool</returns>
        private bool IsOperatorOrParenthesis(char key)
        {
            return operatorFactory.IsValidOperator(key) || this.IsLeftParenthesis(key) || this.IsRightParenthesis(key);
        }

        /// <summary>
        /// Checks if character is a left parenthesis
        /// </summary>
        /// <param name="key">Character to check</param>
        /// <returns>bool</returns>
        private bool IsLeftParenthesis(char key)
        {
            return key == '(';
        }

        /// <summary>
        /// Checks if character is a right parenthesis
        /// </summary>
        /// <param name="key">Character to check</param>
        /// <returns>bool</returns>
        private bool IsRightParenthesis(char key)
        {
            return key == ')';
        }
    }
}
