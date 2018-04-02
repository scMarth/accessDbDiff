using System;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Collections;


namespace accessDbDiff
{
    class Program
    {


        static void Main(string[] args)
        {
            /*
            string accDbPath1 = args[0];
            string accDbPath2 = args[1];

            // Check command line arguments
            if (args.Length != 2) usageExit();
            */

            string basePath = @"C:\localTesting\";
            string outPath = basePath + @"diffResults.txt";
            string accDbPath1 = basePath + @"db1.accdb";
            string accDbPath2 = basePath + @"db2.accdb";

            OleDbConnection connection1 = new OleDbConnection();
            OleDbConnection connection2 = new OleDbConnection();

            connection1.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;"
                + @"Data Source=" + accDbPath1 + ";";

            connection2.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;"
                + @"Data Source=" + accDbPath2 + ";";

            bool connection1Opened = false;
            bool connection2Opened = false;
            bool differenceFound = false;
            bool tryFailed = false;

            try
            {
                connection1.Open();
                connection1Opened = true;
                connection2.Open();
                connection2Opened = true;
                Console.WriteLine("Connected to databases.");

                DataSet ds1 = new DataSet();
                DataSet ds2 = new DataSet();

                // Set the [common] main table name for the databases
                string mainTableName = "TABLE";

                // Query everything within the target tables, order by ascending
                OleDbCommand cmd1 = new OleDbCommand("select * from [" + mainTableName + "] order by [ID] asc", connection1);
                OleDbCommand cmd2 = new OleDbCommand("select * from [" + mainTableName + "] order by [ID] asc", connection2);

                OleDbDataAdapter da1 = new OleDbDataAdapter(cmd1);
                OleDbDataAdapter da2 = new OleDbDataAdapter(cmd2);

                // Populate the datasets
                Console.WriteLine("Populating datasets...");
                da1.Fill(ds1);
                da2.Fill(ds2);
                Console.WriteLine("Datasets populated.");

                // Open the output
                auxFileWriterLib.createFile(outPath);
                // Write a header
                auxFileWriterLib.appendToFile(outPath, "Comparing Files:" + Environment.NewLine);
                auxFileWriterLib.appendToFile(outPath, Environment.NewLine);
                auxFileWriterLib.appendToFile(outPath, "\t" + accDbPath1 + Environment.NewLine);
                auxFileWriterLib.appendToFile(outPath, "\t" + accDbPath2 + Environment.NewLine);
                auxFileWriterLib.appendToFile(outPath, Environment.NewLine);
                auxFileWriterLib.appendToFile(outPath, "========================================================================" + Environment.NewLine);
                auxFileWriterLib.appendToFile(outPath, Environment.NewLine);

                for (int i=0; i<ds1.Tables.Count; i++)
                {
                    DataTable table1 = ds1.Tables[i];
                    DataTable table2 = ds2.Tables[i];

                    // Compare the column headers
                    if (table1.Columns.Count != table2.Columns.Count)
                    {
                        auxFileWriterLib.appendToFile(outPath, "Table " + i + ": Unequal number of columns, skipping." + Environment.NewLine);
                        auxFileWriterLib.appendToFile(outPath, Environment.NewLine);
                        continue;
                    }
                    else
                    {
                        for (int columnIndex=0; columnIndex<table1.Columns.Count; columnIndex++)
                        {
                            string columnCell1 = table1.Columns[columnIndex].ToString();
                            string columnCell2 = table2.Columns[columnIndex].ToString();
                            if (columnCell1 != columnCell2)
                            {
                                auxFileWriterLib.appendToFile(outPath, "Column " + columnIndex + " headers differ between databases:" + Environment.NewLine);
                                auxFileWriterLib.appendToFile(outPath, "\tFile 1: " + columnCell1 + Environment.NewLine);
                                auxFileWriterLib.appendToFile(outPath, "\tFile 2: " + columnCell2 + Environment.NewLine);
                                auxFileWriterLib.appendToFile(outPath, Environment.NewLine);
                            }
                        }
                    }

                    if (table1.Rows.Count != table2.Rows.Count)
                    {
                        differenceFound = true;
                        //Console.WriteLine("Table {0}: Unequal number of rows, skipping.", i);
                        auxFileWriterLib.appendToFile(outPath, "Table " + i + ": Unequal number of rows, skipping." + Environment.NewLine);
                        continue;
                    } else if (table1.Columns.Count != table2.Columns.Count) {
                        differenceFound = true;
                        //Console.WriteLine("Table {0}: Unequal number of columns, skipping.", i);
                        auxFileWriterLib.appendToFile(outPath, "Table " + i + ": Unequal number of columns, skipping." + Environment.NewLine);
                        continue;
                    } else {
                        for (int j = 0; j < table1.Rows.Count; j++)
                        {
                            
                            DataRow row1 = table1.Rows[j];
                            DataRow row2 = table2.Rows[j];

                            // Test if the IDs are the same
                            if (row1[0].ToString() != row2[0].ToString()) // Does an ordinal comparison anyway...
                            //if (row1[0].ToString().CompareTo(row2[0].ToString()) != 0)
                            {
                                differenceFound = true;
                                //Console.WriteLine("Table {0}, Row {1}:", i, j);
                                //Console.WriteLine("\tFile 1 ID: {0}", row1[0]);
                                //Console.WriteLine("\tFile 2 ID: {0}", row2[0]);

                                auxFileWriterLib.appendToFile(outPath, "Table " + i + ", Row " + j + ":" + Environment.NewLine);
                                auxFileWriterLib.appendToFile(outPath, "\tFile 1 ID: " + row1[0] + Environment.NewLine);
                                auxFileWriterLib.appendToFile(outPath, "\tFile 2 ID: " + row2[0] + Environment.NewLine);
                                auxFileWriterLib.appendToFile(outPath, Environment.NewLine);
                                continue;
                            }

                            string ID = row1[0].ToString();

                            for (int k=0; k<table1.Columns.Count; k++)
                            {
                                string cell1 = row1[k].ToString();
                                string cell2 = row2[k].ToString();

                                // Compare strings
                                if (cell1 != cell2) // Does an ordinal comparison anyway...
                                //if (cell1.CompareTo(cell2) != 0)
                                {
                                    differenceFound = true;
                                    //Console.WriteLine("Table {0}, Row {1}, Column {2}:", i, j, k);
                                    //Console.WriteLine("\tFile 1 Value: {0}", cell1);
                                    //Console.WriteLine("\tFile 2 Value: {0}", cell2);

                                    auxFileWriterLib.appendToFile(outPath, "Table " + i + ", ID: " + ID + ", Column " + table1.Columns[k].ToString() + ":" + Environment.NewLine);
                                    auxFileWriterLib.appendToFile(outPath, "\tFile 1 Value: " + cell1 + Environment.NewLine);
                                    auxFileWriterLib.appendToFile(outPath, "\tFile 2 Value: " + cell2 + Environment.NewLine);
                                    auxFileWriterLib.appendToFile(outPath, Environment.NewLine);
                                    continue;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                tryFailed = true;
                if (!connection1Opened)
                {
                    Console.WriteLine("Failed to connect to Access Database 1.");

                    //auxFileWriterLib.appendToFile(outPath, "Failed to connect to Access Database 1.");
                }
                else if (!connection2Opened)
                {
                    Console.WriteLine("Failed to connect to Access Database 2.");

                    //auxFileWriterLib.appendToFile(outPath, "Failed to connect to Access Database 2.");
                }
            }
            finally
            {
                connection1.Close();
                connection2.Close();
            }
            Console.WriteLine();
            if (differenceFound == false && tryFailed == false)
            {
                Console.WriteLine("No differences were found.");
            }
            if (differenceFound == true)
            {
                Console.WriteLine("Differences were found and written to: " + outPath);
            }
            Console.WriteLine();
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        private static void usageExit()
        {
            Console.WriteLine("Usage: accessDbDiff accessDbPath1 accessDbPath2");
            Console.WriteLine("Aborting.");
            Console.ReadKey();
            System.Environment.Exit(1);
        }
    }
}
