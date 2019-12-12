using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

 namespace KartApp
{
    public class Yardimci
    {
        
        public static OleDbConnection Baglanti()
        {
            OleDbConnection olCon = new OleDbConnection();
            olCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= Data.mdb";
            return olCon;
        }
        public static DataTable Tablo(string sorgu)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(sorgu, Baglanti());
            DataTable dt = new DataTable();

            da.Fill(dt);
            return dt;
        }

        public static string Sil(string satir)
        {
            string update = "delete from  AnaTablo where ID= " + satir + "";
            return update;
        }

        public static string SilineniGetir()
        {
            string sorgu = "select ID,Ad,Soyad,Unvan,Tarih,Tel,Gsm,Fax,Mail,Adres,SirketAd,Web from AnaTablo where AktifPasif='0'";
            return sorgu;
        }
        public static string SilKategoriGetir(int kategori)
        {
            string sorgu = "select KatID from AnaTablo where KatID="+kategori+"";
            return sorgu;
        }

        public static string VeriGetir()
        { 
            string sorgu = "SELECT ID,Ad,Soyad,Unvan,Tarih,Tel,Gsm,Fax,Mail,Adres,SirketAd,Web,Kategori.KatAdi From AnaTablo LEFT JOIN Kategori ON AnaTablo.KatID = Kategori.KatID where AktifPasif ='1'";
            return sorgu;
        }
        public static string TumVeriGetir()
        {
            string sorgu = "select ID,Ad,Soyad,Unvan,Tarih,Tel,Gsm,Fax,Mail,Adres,SirketAd,Web from AnaTablo";
            return sorgu;
        }
        public static string VeriKaydet(string Ad,string Soyad,string Unvan,string Tarih,string Tel,string Gsm, string Fax,string Mail,string Adres,string SirketAd, string Web,string Kategori)
        {
            string sorgu = "insert into AnaTablo (Ad, Soyad, Unvan, Tarih, Tel, Gsm, Fax, Mail, Adres, SirketAd, Web, AktifPasif,KatID) values('" + Ad + "', '" + Soyad + "', '" + Unvan + "', '" +Tarih + "', '" + Tel + "','"+Gsm+"', '" + Fax+ "', '" + Mail + "', '" + Adres + "', '" + SirketAd + "', '" + Web + "', '1',"+Kategori+")";
            return sorgu;
        }
        
        public static string VeriGüncelle(string Ad, string Soyad, string Unvan, string Tarih, string Tel,string Gsm, string Fax, string Mail, string Adres, string SirketAd, string Web,string AktifPasif, string Kategori, string satir)
        {
            string sorgu= "update AnaTablo Set Ad= '" + Ad + "',Soyad= '" + Soyad + "',Unvan= '" + Unvan + "',Tarih= '" +Tarih + "',Tel= '" + Tel + "',Gsm='"+Gsm+"',Fax= '" + Fax+ "',Mail= '" + Mail + "',Adres= '" + Adres + "',SirketAd= '" + SirketAd + "',Web= '" + Web + "',AktifPasif='"+AktifPasif+"',KatID="+Kategori+" where ID="+satir+"";
            return sorgu;
        }
        public static string VeriGeriAl(string satir)
        {
            string sorgu= "update AnaTablo set AktifPasif=1 where ID= " + satir + "";
            return sorgu;

        }
        public static string KategoriGetir()
        {
            string sorgu = "select KatID,KatAdi from Kategori";
            return sorgu;
        }
        public static string KategoriKaydet(string KatAdi)
        {
            string sorgu = "insert into Kategori (KatAdi) values('" + KatAdi + "')";
            return sorgu;
        }
        public static string KategoriGüncelle(string KatAdi,string satir)
        {
            string sorgu = "Update Kategori Set KatAdi='" + KatAdi + "' where KatID="+satir+"";
            return sorgu;
        }

    }
}
