using System;

namespace ExpressionTreeDemo
{
    /// <summary>
    /// Demo of Expression Tree DLL
    /// </summary>
    class Program
    {
        /// <summary>
        /// Initializes new ExpressionTree with default expression
        /// </summary>
        private SpreadsheetEngine.ExpressionTree userExpressionTree = new SpreadsheetEngine.ExpressionTree("A1+B1+C1");

        /// <summary>
        /// Runs the ExpressionTree Demo
        /// </summary>
        public void MainMenu()
        {
            // Local Variables
            string userExpression = "A1+B1+C1";
            string userInput = "";
            string varName = "";
            double varValue = 0;

            do
            {
                // Show Menu
                Console.WriteLine("Menu (current expression = \"" + userExpression + "\")");
                Console.WriteLine("1. Enter a new expression");
                Console.WriteLine("2. Set a variable value");
                Console.WriteLine("3. Evaluate tree");
                Console.WriteLine("4. Quit");

                // Get User Option
                userInput = Console.ReadLine().ToString();


                switch (userInput)
                {
                    case "1":
                        Console.WriteLine("Enter a new expression: ");
                        userExpression = Console.ReadLine();
                        userExpressionTree = new SpreadsheetEngine.ExpressionTree(userExpression);                       
                        break;

                    case "2":
                        Console.WriteLine("Enter a variable name: ");
                        varName = Console.ReadLine();
                        Console.WriteLine("Enter a variable value: ");
                        varValue = Convert.ToDouble(Console.ReadLine());
                        userExpressionTree.SetVariable(varName, varValue);
                        break;

                    case "3":
                        Console.WriteLine(userExpressionTree.Evaluate());
                        break;

                    default:
                        break;
                }

            } while (userInput != "4"); // Exit if User enters "4"           
        }

        // Run the Program
        static void Main(string[] args)
        {
            Program ExpressionTreeDemo = new Program();
            ExpressionTreeDemo.MainMenu();
        }
    }
}
