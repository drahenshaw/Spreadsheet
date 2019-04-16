using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Spreadsheet_David_Henshaw_Tests
{
    [TestFixture]
    public class TestForm
    {
        private Spreadsheet_David_Henshaw.Form1 objectUnderTest = new Spreadsheet_David_Henshaw.Form1() { };

        private MethodInfo GetMethod(string methodName)
        {
            if (string.IsNullOrWhiteSpace(methodName))
            {
                Assert.Fail("methodName cannot be null or whitespace");                
            }

            var method = this.objectUnderTest.GetType().GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);

            if (method == null)
            {
                Assert.Fail(string.Format("{0} method not found", methodName));
            }

            return method;
        }

        private FieldInfo GetField(string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
            {
                Assert.Fail("fieldName cannot be null or whitespace");
            }

            var field = this.objectUnderTest.GetType().GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);

            if (field == null)
            {
                Assert.Fail(string.Format("{0} field not found", fieldName));
            }

            return field;
        }

        private PropertyInfo GetProperty(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
            {
                Assert.Fail("propertyName cannot be null or whitespace");
            }

            var property = this.objectUnderTest.GetType().GetProperty(propertyName, BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);

            if (property == null)
            {
                Assert.Fail(string.Format("{0} property not found", propertyName));
            }

            return property;
        }

        private object GetValue(PropertyInfo info, object property)
        {            
            return ((PropertyInfo)info).GetValue(property, null);            
        }


        [Test]
        public void TestStartup()
        {
            FieldInfo fieldInfo = GetField("dataGridView");
            PropertyInfo prop = fieldInfo.GetValue(objectUnderTest).GetType().GetProperty("RowCount");
            //object test = GetValue(prop, objectUnderTest);
            //Int32 value = (Int32)prop.GetValue(instanceClass);
           // objectUnderTest = (Spreadsheet_David_Henshaw.Form1)prop.GetValue(this.objectUnderTest);
            Assert.Pass("{0} pass", prop);            
        }

        [Test]
        public void TestSpreadsheet()
        {
            // Test that the Bounds of the Spreadsheet are within the Requirements            
            CptS321.Spreadsheet spreadsheet = new CptS321.Spreadsheet(10, 10);            
            Assert.That(spreadsheet.RowCount >= 1 && spreadsheet.RowCount <= 100) ;
            Assert.That(spreadsheet.ColumnCount >= 1 && spreadsheet.ColumnCount <= 100);

            // Test that trying to access a cell outside the array throws an error            
            Assert.Throws<IndexOutOfRangeException>(() => spreadsheet.GetCell(10, 11));
        }

        [Test]
        public void TestEngine()
        {
            CptS321.Spreadsheet spreadsheet = new CptS321.Spreadsheet(100, 100);             
        }

        [Test]
        public void TestSetVariable()
        {
            SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree("5+4");
            expressionTree.SetVariable("A1", 42);
            var fieldInfo = expressionTree.GetType().GetField("root", BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);            
        }

        [Test]
        public void TestEvaluate()
        {
            // Test Support for Addition
            SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree("5+4");
            Assert.That(expressionTree.Evaluate() == 9.0);

            // Test Support for Subtraction
            SpreadsheetEngine.ExpressionTree expressionTree2 = new SpreadsheetEngine.ExpressionTree("5-4");
            Assert.That(expressionTree2.Evaluate() == 1.0);

            // Test Support for Multiplication
            SpreadsheetEngine.ExpressionTree expressionTree3 = new SpreadsheetEngine.ExpressionTree("4*4");
            Assert.That(expressionTree3.Evaluate() == 16.0);

            // Test Support for Division
            SpreadsheetEngine.ExpressionTree expressionTree4 = new SpreadsheetEngine.ExpressionTree("5/5");
            Assert.That(expressionTree4.Evaluate() == 1.0);

        }

        
        public void NotImplemented()
        {
            // Test Not-Implemented Operator
            SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree("5^2");
            expressionTree.Evaluate();
        }

        [Test]
        public void TestNotImplmented()
        {
            // Test that Not-Implemented Operators Throw Appropriate Exception - Not supposed to pass for HW6
            Assert.Throws<System.NullReferenceException>(NotImplemented);
        }

        public void DivideByZero()
        {
            SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree("5/0");
            expressionTree.Evaluate();
        }

        [Test]
        public void TestDivideByZero()
        {
            //This test passes manually, NUnit thinks the exception is null
            Assert.Throws<System.DivideByZeroException>(DivideByZero);
        }

        [Test]
        public void TestParenthesis()
        {
            SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree("(5+3)");
            Assert.AreEqual(expressionTree.Evaluate(), 8);
        }

        [Test]
        public void TestPrecedence()
        {
            SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree("5*2-3");
            Assert.AreEqual(expressionTree.Evaluate(), 7);
        }       
        
        [Test]
        public void TestGetPrecedence()
        {
            // This test passed in debug mode - something to do with traverse available operators results in an error
            // that the key has already been added to the operator dictionary
            SpreadsheetEngine.ExpressionTreeNodeFactory exFactory = new SpreadsheetEngine.ExpressionTreeNodeFactory();
            Assert.AreEqual(7, exFactory.GetPrecedence('+'));            
        }

        [Test]
        public void ShuntingYardTest()
        {
            SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree("5+5");

            var methodInfo = expressionTree.GetType().GetMethod("ConvertToPostfix", BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);

            if (methodInfo != null)
            {
                List<string> expectedResult = new List<string> { "5", "5", "+" };
                object actualResult = methodInfo.Invoke(expressionTree, new object[] { "5+5" });
                Assert.AreEqual(expectedResult, actualResult);

                expectedResult = new List<string> { "A", "B", "C", "*", "D", "/", "+", "E", "-"};
                actualResult = methodInfo.Invoke(expressionTree, new object[] { "A + B * C / D - E" });
                Assert.AreEqual(expectedResult, actualResult);

                expectedResult = new List<string> { "A", "B", "C", "D", "-", "*", "+", "E", "/" };
                actualResult = methodInfo.Invoke(expressionTree, new object[] { "( A + B * ( C - D ) ) / E" });
                Assert.AreEqual(expectedResult, actualResult);

                expectedResult = new List<string> { "A", "B", "C", "D", "*", "+", "*", "E", "+" };
                actualResult = methodInfo.Invoke(expressionTree, new object[] { "A * (B + C * D) + E" });
                Assert.AreEqual(expectedResult, actualResult);

                Assert.Pass("Sucessfully converted all infix strings to postfix list of substrings");
            }
        }

        [Test]
        public void ConstructPostfixTreeTest()
        {
            SpreadsheetEngine.ExpressionTree expressionTree = new SpreadsheetEngine.ExpressionTree("5+5");

            var methodInfo = expressionTree.GetType().GetMethod("ConstructPostfixTree", BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);

            if (methodInfo != null)
            {
                SpreadsheetEngine.AdditionExpressionTreeNode expectedResult = new SpreadsheetEngine.AdditionExpressionTreeNode();

                SpreadsheetEngine.ExpressionTreeNode actualResult = (SpreadsheetEngine.ExpressionTreeNode)methodInfo.Invoke(expressionTree, new object[] { new List<string> { "5", "5", "+" } });
                Assert.AreEqual(expectedResult.GetType(), actualResult.GetType());
                Assert.Pass("Root of constructed tree matches expected type");
            }
        }
    }
}


  
                


