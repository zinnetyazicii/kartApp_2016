
using System;
using System.IO;
using System.Drawing;
using System.Drawing.Printing;

class Printer
{
    private static StreamReader dosyaAkimi;

    static void Main(string[] args)
    {
        dosyaAkimi = new System.IO.StreamReader("C:\\Data.mdb");
        
        PrintDocument PD = new PrintDocument();
        PD.PrintPage += new PrintPageEventHandler(OnPrintDocument);

        try
        {
            PD.Print();
        }
        catch
        {
            Console.WriteLine("Yazici çiktisi alinamiyor...");
        }
        finally
        {
            PD.Dispose();
        }
    }

    public static void OnPrintDocument(object sender, PrintPageEventArgs e)
    {
        Font font = new Font("Verdana", 11);
        float yPozisyon = 0; int LineCount = 0;
        float leftMargin = e.MarginBounds.Left;
        float topMargin = e.MarginBounds.Top;

        string line = null;

        float SayfaBasinaDusenSatir = e.MarginBounds.Height / font.GetHeight();

        while (((line = dosyaAkimi.ReadLine()) != null) && LineCount < SayfaBasinaDusenSatir)
        {
            yPozisyon = topMargin + (LineCount * font.GetHeight(e.Graphics));
            e.Graphics.DrawString(line, font, Brushes.Red, leftMargin, yPozisyon);

            LineCount++;
        }

        if (line == null)
            e.HasMorePages = false;
        else
            e.HasMorePages = true;

    }
}