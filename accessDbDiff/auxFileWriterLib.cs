class auxFileWriterLib
{
    // Create an empty file 'outfile'. Overwrite if it already exists
    public static void createFile(string outfile)
    {
        System.IO.File.WriteAllText(outfile, "");
    }

    // Append string 'str' to output CSV file 'filename'
    public static void appendToFile(string filename, string str)
    {
        using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(filename, true))
        {
            file.Write(str);
            file.Flush();
            file.Close();
        }
    }
}

