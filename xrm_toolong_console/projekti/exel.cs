
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;

namespace projekti
{
    class exel
    {
        public static string Customer_Name;
        public static string Address;
        public static string Email;
        public static string Description;
        public static string Number;
        public static string Total;
        public static ArrayList Product;
        public static ArrayList Product_Code;
        public static ArrayList Price;
        public static ArrayList Unit;
        public static ArrayList Quantity;
        public static ArrayList Amount;

        public static void setCustomer_Name(string x)
        {
            Customer_Name = x;
        }
        public static void setAddress(string x)
        {
            Address = x;
        }
        public static void setEmail(string x)
        {
            Email = x;
        }
        public static void setDescription(string x)
        {
            Description = x;
        }
        public static void setNumber(string x)
        {
            Number = x;
        }
        public static void setTotal(string x)
        {
            Total = x;
        }
        public static void setProduct(ArrayList x)
        {
            Product = x;
        }
        public static void setProduct_Code(ArrayList x)
        {
            Product_Code = x;
        }
        public static void setPrice(ArrayList x)
        {
            Price = x;
        }
        public static void setUnit(ArrayList x)
        {
            Unit = x;
        }
        public static void setQuantity(ArrayList x)
        {
            Quantity = x;
        }
        public static void setAmount(ArrayList x)
        {
            Amount = x;
        }

        internal static void FindSpecificChar(object[,] valueArray, int rowCount, int colCount)
        {
            int nr = 0;
            int colInExel = 1;
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (!rowIsNull(valueArray, i, colCount)) //timesaver
                    {
                        if (!(valueArray[i, j] is null) && !(valueArray[i + 1, j] is null))
                        {

                            if (valueArray[i, j].ToString().Contains("Customer Name"))
                            {
                                setCustomer_Name(valueArray[i + 1, j].ToString());
                            }
                            else if (valueArray[i, j].ToString().Contains("Address"))
                            {
                                setAddress(valueArray[i + 1, j].ToString());
                            }
                            else if (valueArray[i, j].ToString().Contains("Email"))
                            {
                                setEmail(valueArray[i + 1, j].ToString());
                            }
                            else if (valueArray[i, j].ToString().Contains("Description"))
                            {
                                setDescription(valueArray[i + 1, j].ToString());
                            }
                            else if (valueArray[i, j].ToString().Contains("Number"))
                            {
                                setNumber(valueArray[i + 1, j].ToString());
                            }
                            else if (valueArray[i, j].ToString().Contains("Total"))
                            {
                                setTotal(valueArray[i + 1, j].ToString());
                            }
                            else if ((valueArray[i, j].ToString().Contains("Product")) && !(valueArray[i, j].ToString().Contains("Code")))
                            {
                                ArrayList arlist = new ArrayList();
                                int productNumber = findNrProducts(valueArray, i, j);
                                nr = productNumber;
                                for (int f = 1; f <= productNumber; f++)
                                {
                                    arlist.Add(valueArray[i + f, j]);
                                }
                                setProduct(arlist);
                            }
                            else if (valueArray[i, j].ToString().Contains("Product Code"))
                            {
                                ArrayList arlist = new ArrayList();

                                for (int f = 1; f <= nr; f++)
                                {
                                    arlist.Add(valueArray[i + f, j]);
                                }
                                setProduct_Code(arlist);
                            }
                            else if (valueArray[i, j].ToString().Contains("Price"))
                            {
                                ArrayList arlist = new ArrayList();

                                for (int f = 1; f <= nr; f++)
                                {
                                    arlist.Add(valueArray[i + f, j]);
                                }
                                setPrice(arlist);
                            }
                            else if (valueArray[i, j].ToString().Contains("Unit"))
                            {
                                ArrayList arlist = new ArrayList();

                                for (int f = 1; f <= nr; f++)
                                {
                                    arlist.Add(valueArray[i + f, j]);
                                }
                                setUnit(arlist);
                            }
                            else if (valueArray[i, j].ToString().Contains("Quantity"))
                            {
                                ArrayList arlist = new ArrayList();

                                for (int f = 1; f <= nr; f++)
                                {
                                    arlist.Add(valueArray[i + f, j]);
                                }
                                setQuantity(arlist);
                            }
                            else if (valueArray[i, j].ToString().Contains("Amount"))
                            {
                                ArrayList arlist = new ArrayList();

                                for (int f = 1; f <= nr; f++)
                                {
                                    arlist.Add(valueArray[i + f, j]);
                                }
                                setAmount(arlist);
                            }
                        }
                    }
                }

                if (rowIsNull(valueArray, i, colCount))
                {
                    colInExel++;

                    if (colInExel > 40) //timesaver :only 40 rows
                    {
                        break;
                    }
                }

            }

        }

        private static int findNrProducts(object[,] valueArray, int i, int j)
        {
            int nr = 0;
            int k = j;
            while (1 == 1)
            {
                if (!(valueArray[i + 1, j] is null))
                {
                    i++;
                    nr++;
                }
                else
                {
                    break;
                }
            }
            return nr;
        }

        private static bool rowIsNull(object[,] valueArray, int i, int colCount)
        {
            for (int j = 1; j <= colCount; j++)
            {
                if (!(valueArray[i, j] is null))
                {
                    return false;
                }
            }
            return true;
        }
    }
}