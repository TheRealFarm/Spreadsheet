/* 
 * Author: Kyle Farmer
 * ID: 11342935
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CptS322
{
    public class ExpTree
    {
        public abstract class Node
        {
            internal Node L, R;
        }

        // Node for operators within the ExpTree
        public class OpNode : Node
        {
            public char m_operand;

            public OpNode(char op)
            {
                m_operand = op;
            }
        }

        // Node for variables in ExpTree
        public class VarNode : Node
        {
            public string m_var;

            public VarNode(string var)
            {
                m_var = var;
            }
        }

        // Node for constants in ExpTree
        public class ConstantNode : Node
        {
            public double m_constant;

            public ConstantNode(double num)
            {
                m_constant = num;
            }
        }

        private Node m_root; // root node for expression tree
        private string m_expression;
        public readonly static char[] PossibleOperators = { '+', '-', '*', '/' };
        public Dictionary<string, double> VarDictionary = new Dictionary<string, double>();

        // getter for root
        private Node RootProperty
        {
            get { return this.m_root; }
            //set { this.m_root = value; }
        }


        public ExpTree(string expression)
        {
            // set the internal string, clear the variable dictionary, and recompile tree
            // when a new string is used to construct the tree
            m_expression = expression;
            VarDictionary.Clear();
            m_root = Compile(m_expression);
        }
        // getter/setter for expression

        public double Eval()
        {
            return Eval(m_root);
        }

        // private Eval helper function for public Eval()
        // takes input node argument and evaluates it
        private double Eval(Node NodeArg)
        {
            ConstantNode cNode = NodeArg as ConstantNode;

            if (cNode != null)
            {
                return cNode.m_constant;
            }

            VarNode vNode = NodeArg as VarNode;
            if (vNode != null)
            {
                return VarDictionary[vNode.m_var];
            }

            // if node is operator, recursively evaluate left and right subtrees
            // and perform respective operations on them
            OpNode oNode = NodeArg as OpNode;
            if (oNode != null)
            {
                switch (oNode.m_operand) // evalutate the left subtree and then right subtree
                {
                    case '+':
                        return Eval(oNode.L) + Eval(oNode.R);

                    case '-':
                        return Eval(oNode.L) - Eval(oNode.R);

                    case '*':
                        return Eval(oNode.L) * Eval(oNode.R);

                    case '/':
                        return Eval(oNode.L) / Eval(oNode.R);
                }
            }
            return 0;
        }

        // Compile will use the new expression to create an expression tree for evaluation
        // Returns null if the new string is empty
        private Node Compile(string newExpression)
        {
            if (string.IsNullOrEmpty(newExpression))
                return null;

            //Detects left paranthesis, begins looking for right parenthesis
            if (newExpression[0] == '(')
            {
                // Counter to keep track of parenthesis
                int pCounter = 0;
                for (int i = 0; i < newExpression.Length; i++)
                {
                    //Increments at left parenthesis
                    if (newExpression[i] == '(')
                    {
                        pCounter++;
                    }

                    //Decrements at right parenthesis
                    else if (newExpression[i] == ')')
                    {
                        pCounter--;

                        if (pCounter == 0)//Counter at zero means left and right parenthesis have matched
                        {
                            if (newExpression.Length - 1 != i)
                            {
                                break; // if we are not at string end, we continue compilation
                            }
                            else
                            {
                                // if we are at end of expression, we compile between the parenthesis
                                return Compile(newExpression.Substring(1, newExpression.Length - 2));
                            }
                        }
                    }
                }
            }


            char[] OpArray = ExpTree.PossibleOperators;
            foreach (char operation in OpArray)
            {
                // compile the expression based on the current operation
                // only return subtree if node is non-null
                Node oNode = Compile(newExpression, operation);
                if (oNode != null)
                    return oNode;
            }

            // node is either constant or variable
            double num;
            if (double.TryParse(newExpression, out num))
            {
                return new ConstantNode(num);
            }
            else
            {
                // initiliaze the variable in the dcitionary when found
                VarDictionary[newExpression] = 0;
                return new VarNode(newExpression);
            }
        }

        // This is a compile utility function that will serve as a recursive call for compile
        // it will take an expression and an operator to return a branch node of the expression
        private Node Compile(string expression, char operation)
        {
            bool flag = false;
            // we want to read from the right to find the rightmost lowest precedence operator
            int i = expression.Length - 1;
            int pCounter = 0;

            while (!flag)
            {
                // increment parenthesis counter
                if (expression[i] == '(')
                {
                    pCounter++;
                }
                // decrement parenthesis counter
                else if (expression[i] == ')')
                {
                    pCounter--;
                }

                if (pCounter == 0 && expression[i] == operation) // when we reach an operator, evaluate expression
                {
                    // create and return subtree with current operation being the root node
                    // and the left and right expressions as their own compiled subtrees
                    OpNode opNode = new OpNode(operation);
                    // left subtree is the beginning of the string to the current index
                    opNode.L = Compile(expression.Substring(0, i));
                    // right subtree is the next index to the end of the string
                    opNode.R = Compile(expression.Substring(i + 1));
                    return opNode;
                }
                else
                {
                    if (i == 0) // reached the left end of the expression - terminate
                        flag = true;
                    i--; // continue reading the next part of the expression to the left
                }
            }
            return null;
        }

        // Function used to set variable values in expression
        public void SetVar(string VarName, double VarValue)
        {
            VarDictionary[VarName] = VarValue;
        }

        // Utility function for spreadsheet class.
        // Gathers all expression variables returned as an array of strings.
        public string[] GetVarNames()
        {
            return VarDictionary.Keys.ToArray();
        }
    }
}

