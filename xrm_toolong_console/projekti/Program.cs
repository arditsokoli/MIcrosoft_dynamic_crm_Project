using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace projekti
{
	class Program
	{
		protected static Guid guid_customer;
		protected static Guid guid_product;
		protected static Guid guid_invoice;
		public static string errorsFilepath = @"c:\FISOFT_Consulting\Error_Logs\errors.txt";

		public static void Setguid_customer(Guid value)
		{
			guid_customer = value;
		}
		public static void Setguid_product(Guid value)
		{
			guid_product = value;
		}
		public static void Setguid_invoice(Guid value)
		{
			guid_invoice = value;
		}

		static void Main(string[] args)
		{
			File.WriteAllText(errorsFilepath, " "); //clean last messages in error_file
			try
			{
				//Connect to server:
				string connectionString = @"AuthType=OAuth;
                                        Username=DitiSokoli@fisoft989.onmicrosoft.com;
                                        Password=Password1;
                                        Url=https://org85c04817.crm4.dynamics.com/;
                                        AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;
                                        RedirectUri=app://58145B91-0C36-4500-8554-080854F2AC97";
				CrmServiceClient service = new CrmServiceClient(connectionString);

				String queryCustomer = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='fisoft_customer'>
                                    <attribute name='fisoft_customerid'/>
                                    <attribute name='fisoft_name'/>
                                    <attribute name='createdon'/>
                                    <order attribute='fisoft_name' descending='false'/>
                                    <filter type='and'>
                                      <condition attribute='statecode' operator='eq' value='0'/>
                                    </filter>
                                  </entity>
                                </fetch>";
				EntityCollection collectionCustomer = service.RetrieveMultiple(new FetchExpression(queryCustomer));


				String queryProduct = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='fisoft_product'>
                                    <attribute name='fisoft_productid'/>
                                    <attribute name='fisoft_name'/>
                                    <attribute name='createdon'/>
                                    <order attribute='fisoft_name' descending='false'/>
                                    <filter type='and'>
                                      <condition attribute='statecode' operator='eq' value='0'/>
                                    </filter>
                                  </entity>
                                </fetch>";
				EntityCollection collectionProduct = service.RetrieveMultiple(new FetchExpression(queryProduct));

				String queryInvoice = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='fisoft_invoice'>
                                    <attribute name='fisoft_invoiceid' />
                                    <attribute name='fisoft_name' />
                                    <attribute name='createdon' />
                                    <attribute name='fisoft_total_amount' />
                                    <attribute name='fisoft_paid' />
                                    <order attribute='fisoft_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>";
				EntityCollection collectionInvoice = service.RetrieveMultiple(new FetchExpression(queryInvoice));

				Application excelApp = new Application();
				if (excelApp != null)
				{
					Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\FISOFT_Consulting\invoices\Invoice Demo.xlsx");
					Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];

					Range excelRange = excelWorksheet.UsedRange;
					int rowCount = excelRange.Rows.Count;
					int colCount = excelRange.Columns.Count;

					object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
					exel.FindSpecificChar(valueArray, rowCount, colCount);

					DateTime todayDate = DateTime.UtcNow.Date;
					bool customerIsnew = false;
					bool productIsnew = false;


					if (!Exist(exel.Customer_Name, queryCustomer, collectionCustomer, "fisoft_name"))
					{
						customerIsnew = true;
						File.AppendAllText(errorsFilepath, Environment.NewLine + "Creating new Customer." + Environment.NewLine);
						Entity customer = new Entity("fisoft_customer");
						customer.Attributes.Add("fisoft_name", exel.Customer_Name);
						customer.Attributes.Add("emailaddress", exel.Email);
						customer.Attributes.Add("fisoft_address", exel.Address);
						customer.Attributes.Add("fisoft_revenue", exel.Total.ToString());
						Guid guid_customer = service.Create(customer);
						Setguid_customer(guid_customer);
					}
					if (((!Exist(exel.Customer_Name, queryCustomer, collectionCustomer, "fisoft_name")) ||
							(Exist(exel.Customer_Name, queryCustomer, collectionCustomer, "fisoft_name"))))
					{
						File.AppendAllText(errorsFilepath, Environment.NewLine + "Creating new Invoice." + Environment.NewLine);
						Entity invoice = new Entity("fisoft_invoice");
						invoice.Attributes.Add("fisoft_description", exel.Description);
						invoice.Attributes.Add("fisoft_number", exel.Number);
						invoice.Attributes.Add("fisoft_createddate", todayDate);

						OptionSetValue myOptionSet = new OptionSetValue();
						myOptionSet.Value = 874630001;//set to default "no"	

						invoice.Attributes.Add("fisoft_paid", myOptionSet);
						double total = Double.Parse(exel.Total);
						invoice.Attributes.Add("fisoft_total_amount", total);

						if (customerIsnew)
						{ //lookup filed validation
							invoice.Attributes["fisoft_customer"] = new EntityReference("fisoft_customer", guid_customer);
						}
						else
						{
							foreach (Entity Customer in collectionCustomer.Entities)
							{
								var Guid = new Guid(Customer.Attributes["fisoft_customerid"].ToString());
								invoice.Attributes["fisoft_customer"] = new EntityReference("fisoft_customer", Guid);
							}
						}
						Guid guid_invoice = service.Create(invoice);
						Setguid_invoice(guid_invoice);
					}

					for (int d = 0; d < exel.Product.Count; d++)
					{
						if (!Exist(exel.Product[d].ToString(), queryProduct, collectionProduct, "fisoft_name"))
						{
							productIsnew = true;
							File.AppendAllText(errorsFilepath, Environment.NewLine + "Creating new product." + Environment.NewLine);
							Entity product = new Entity("fisoft_product");
							product.Attributes.Add("fisoft_name", exel.Product[d]);
							double price = Double.Parse(exel.Price[d].ToString());
							product.Attributes.Add("fisoft_price", price);
							product.Attributes.Add("fisoft_product_code", exel.Product_Code[d]);
							Guid guid_product = service.Create(product);
							Setguid_product(guid_product);
						}
						if ((!Exist(exel.Product[d].ToString(), queryProduct, collectionProduct, "fisoft_name")) ||
							(Exist(exel.Product[d].ToString(), queryProduct, collectionProduct, "fisoft_name")))
						{
							File.AppendAllText(errorsFilepath, Environment.NewLine + "Creating new invoice Product." + Environment.NewLine);
							Entity invoice_product = new Entity("fisoft_inoice_product");

							if (productIsnew)
							{
								invoice_product.Attributes["fisoft_product"] = new EntityReference("fisoft_product", guid_product);
							}
							else
							{
								foreach (Entity Product in collectionProduct.Entities)
								{
									var Guid = new Guid(Product.Attributes["fisoft_productid"].ToString());
									invoice_product.Attributes["fisoft_product"] = new EntityReference("fisoft_product", Guid);
								}
							}
							invoice_product.Attributes["fisoft_invoice"] = new EntityReference("fisoft_invoice", guid_invoice);
							invoice_product.Attributes.Add("fisoft_unit", exel.Unit[d]);
							invoice_product.Attributes.Add("fisoft_quantity", exel.Quantity[d]);
							invoice_product.Attributes.Add("fisoft_amount", exel.Amount[d]);
							Guid guid_invoice_product = service.Create(invoice_product);
						}
					}



					//finding totalAmount
					var totalAmount = "";
					double revenue = Double.Parse(exel.Total);
					foreach (Entity x in collectionInvoice.Entities)
					{
						totalAmount = x.Attributes["fisoft_total_amount"].ToString();
						revenue += Double.Parse(totalAmount);
					}

					//update
					if (revenue != Double.Parse(exel.Total))
					{
						Entity customer = new Entity("fisoft_customer");
						foreach (Entity customeri in collectionCustomer.Entities)
						{
							var id = new Guid(customeri.Attributes["fisoft_customerid"].ToString());
							customer.Id = id;
						}

						customer.Attributes["fisoft_revenue"] = revenue.ToString();
						service.Update(customer);
					}
					Console.WriteLine("Success compilation ! \nCheck error file in \\FISOFT_Consulting\\Error_Logs\\errors.txt ");
					excelWorkbook.Close();
					excelApp.Quit();
				}
			}
			catch (Exception e)
			{
				Console.WriteLine("Error code !!! \nCheck error file in \\FISOFT_Consulting\\Error_Logs\\errors.txt ");
				File.AppendAllText(errorsFilepath, Environment.NewLine + "Error: " + e.Message + Environment.NewLine);
			}
		}

		private static bool Exist(string Name, string query, EntityCollection collection, string n)
		{
			foreach (Entity x in collection.Entities)
			{
				if (x.Attributes[n].ToString().Contains(Name))
				{
					return true;
					break;
				}
			}
			return false;
		}
	}
}