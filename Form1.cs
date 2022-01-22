using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Xml;

namespace EtaRapor
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }
        private int _errorplus = 0;
        private string kur = "USD", tur, dolar = string.Empty;
        private string gun, ay;
        private readonly SqlConnection myConnection = new SqlConnection(File.ReadAllText("database.txt"));
        private XmlDocument xml = new XmlDocument();
        void List(string list)
        {
            SqlDataAdapter da = new SqlDataAdapter(list, myConnection);
            DataSet ds = new DataSet();
            da.Fill(ds);
            gridControl1.DataSource = ds.Tables[0];
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dateEdit1.EditValue = "01.01.2022";
            dateEdit2.EditValue = "31.01.2022";
            List("select K.CARKOD as 'CARİ KODU', K.CARUNVAN AS 'CARİ ÜNVAN',H.CARHAREVRAKNO AS 'EVRAK NO',H.CARHARACIKLAMA AS'AÇIKLAMA',CARHARISTIPKOD AS'FİŞ TİPİ',CAST(H.CARHARDOVKUR  as DECIMAL(15,3))AS 'DÖVİZ KURU',H.CARHARDOVTUR AS 'DÖVİZ TÜR',H.CARHARDOVKOD AS 'DÖVİZ KOD',H.CARHARTAR AS 'HAREKET TARİHİ',CAST((H.CARHARTUTAR-H.CARHARKDVTUTAR) as DECIMAL(15,3)) AS 'MAL HİZMET TUTARI',CAST(H.CARHARKDVTUTAR as DECIMAL(15,3)) AS 'KDV TUTARI',CAST(H.CARHARTUTAR as DECIMAL(15,3)) as 'HAREKET TOPLAM',CAST(H.CARHARDOVTUTAR as DECIMAL(15,3)) AS 'DÖVİZ TOPLAM',CAST(K.CARBAKIYE as DECIMAL(15,3)) 'CARİ BAKİYE','" + kur + "' as 'ARANAN DÖVİZ','' AS 'DÖVİZİN KURU' ,'' AS 'DOVİZ KARŞILIGI' from CARKART K  LEFT JOIN CARHAR H  \r\n ON \r\n (K.CARKOD=H.CARHARCARKOD   and H.CARHARTAR>='" + dateEdit1.DateTime.ToString("yyyy-MM-dd") + "'\r\n  and H.CARHARTAR<='" + dateEdit2.DateTime.ToString("yyyy-MM-dd") + "'\r\n ) \r\n where H.CARHARISTIPKOD='FATURA'");
            Lbl_Count.Text = gridView1.RowCount.ToString();
        }
        private void eXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                gridView1.ExportToXlsx(folderBrowserDialog1.SelectedPath + "//CariListe.xlsx");
            }
        }
        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                gridView1.ExportToPdf(folderBrowserDialog1.SelectedPath + "//CariListe.pdf");
            }
        }

        private void gridView1_RowCountChanged(object sender, EventArgs e)
        {
            Lbl_Count.Text = gridView1.RowCount.ToString();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            List("select K.CARKOD as 'CARİ KODU', K.CARUNVAN AS 'CARİ ÜNVAN',H.CARHAREVRAKNO AS 'EVRAK NO',H.CARHARACIKLAMA AS'AÇIKLAMA',CARHARISTIPKOD AS'FİŞ TİPİ',CAST(H.CARHARDOVKUR  as DECIMAL(15,3))AS 'DÖVİZ KURU',H.CARHARDOVTUR AS 'DÖVİZ TÜR',H.CARHARDOVKOD AS 'DÖVİZ KOD',H.CARHARTAR AS 'HAREKET TARİHİ',CAST((H.CARHARTUTAR-H.CARHARKDVTUTAR) as DECIMAL(15,3)) AS 'MAL HİZMET TUTARI',CAST(H.CARHARKDVTUTAR as DECIMAL(15,3)) AS 'KDV TUTARI',CAST(H.CARHARTUTAR as DECIMAL(15,3)) as 'HAREKET TOPLAM',CAST(H.CARHARDOVTUTAR as DECIMAL(15,3)) AS 'DÖVİZ TOPLAM',CAST(K.CARBAKIYE as DECIMAL(15,3)) 'CARİ BAKİYE','" + kur + "' as 'ARANAN DÖVİZ','' AS 'DÖVİZİN KURU' ,'' AS 'DOVİZ KARŞILIGI' from CARKART K  LEFT JOIN CARHAR H  \r\n ON \r\n (K.CARKOD=H.CARHARCARKOD   and H.CARHARTAR>='" + dateEdit1.DateTime.ToString("yyyy-MM-dd") + "'\r\n  and H.CARHARTAR<='" + dateEdit2.DateTime.ToString("yyyy-MM-dd") + "'\r\n ) \r\n where H.CARHARISTIPKOD='FATURA'");
            Lbl_Count.Text = gridView1.RowCount.ToString();
        }
        private void comboBoxEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            kur = comboBoxEdit1.SelectedItem.ToString();
        }
        private void CurrenyList_Click(object sender, EventArgs e)
        {
            for (int i = _errorplus; i < gridView1.RowCount; i++)
            {
                try
                {
                    DateTime day = Convert.ToDateTime(gridView1.GetRowCellValue(i, "HAREKET TARİHİ"));
                    if (day.Day.ToString().Length < 2)
                        gun = "0" + day.Day.ToString();
                    else
                        gun = day.Day.ToString();
                    if (day.Month.ToString().Length < 2)
                        ay = "0" + day.Month.ToString();
                    else
                        ay = day.Month.ToString();
                    xml.Load("https://www.tcmb.gov.tr/kurlar/" + day.Year.ToString() + "" + ay + "/" + gun + "" + ay +
                             "" + day.Year.ToString() + ".xml");
                    tur = gridView1.GetRowCellValue(i, "DÖVİZ TÜR").ToString();
                    if ("AKTALS" == tur)
                        dolar = xml.SelectSingleNode("Tarih_Date/Currency [@Kod='" + kur + "']/BanknoteBuying").InnerXml;
                    else if ("AKTALS" == tur)
                        dolar = xml.SelectSingleNode("Tarih_Date/Currency [@Kod='" + kur + "']/BanknoteSelling").InnerXml;
                    else if ("MBNKALS" == tur)
                        dolar = xml.SelectSingleNode("Tarih_Date/Currency [@Kod='" + kur + "']/ForexBuying").InnerXml;
                    else if ("MBNKSAT" == tur)
                        dolar = xml.SelectSingleNode("Tarih_Date/Currency [@Kod='" + kur + "']/ForexSelling").InnerXml;
                    else
                        dolar = xml.SelectSingleNode("Tarih_Date/Currency [@Kod='" + kur + "']/ForexBuying").InnerXml;
                    gridView1.SetRowCellValue(i, "DÖVİZİN KURU", dolar.Replace('.', ',').Substring(0, 5));
                    gridView1.SetRowCellValue(i, "ARANAN DÖVİZ", kur);
                    decimal deger2 = (((decimal)gridView1.GetRowCellValue(i, "HAREKET TOPLAM")) / Convert.ToDecimal(dolar.Replace('.', ',').Substring(0, 5)));
                    gridView1.SetRowCellValue(i, "DOVİZ KARŞILIGI", Math.Round(deger2, 3));
                    decimal deger = (((decimal)gridView1.GetRowCellValue(i, "DOVİZ KARŞILIGI")) * Convert.ToDecimal(dolar.Replace('.', ',').Substring(0, 5)));
                    _errorplus = i;
                }
                catch (Exception a)
                {
                    _errorplus++;
                }
                continue;
            }
            _errorplus = 0;
        }
    }
}