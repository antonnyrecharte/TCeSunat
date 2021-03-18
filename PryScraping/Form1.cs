using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace PryScraping
{
    public partial class Form1 : Form
    {
        private static IWebDriver driver;
        private ClsETCeSunat pTCeSunat;
        
        public Form1()
        {
            InitializeComponent();
            driver = new ChromeDriver();
        }
        
        private static int pYear;


        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                pTCeSunat = new ClsETCeSunat(DateTime.Today.Year, DateTime.Today.Month);

                cboAnio.SelectedValue = pTCeSunat.TCSanio;
                cboMes.SelectedValue = pTCeSunat.TCSmes;

                LlenarCombo();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al Guardar \n" + ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void LlenarCombo()
        {
            try
            {
                Dictionary<int, string> ListaMes = new Dictionary<int, string>();
                ListaMes.Add(1, "Enero");
                ListaMes.Add(2, "Febrero");
                ListaMes.Add(3, "Marzo");
                ListaMes.Add(4, "Abril");
                ListaMes.Add(5, "Mayo");
                ListaMes.Add(6, "Junio");
                ListaMes.Add(7, "Julio");
                ListaMes.Add(8, "Agosto");
                ListaMes.Add(9, "Septiembre");
                ListaMes.Add(10, "Octubre");
                ListaMes.Add(11, "Noviembre");
                ListaMes.Add(12, "Diciembre");

                cboMes.DisplayMember = "Value";
                cboMes.ValueMember = "Key";
                cboMes.DataSource = ListaMes.ToArray();

                pYear = DateTime.Today.Year;

                Dictionary<int, string> ListaAnio = new Dictionary<int, string>();
                for (int i = pYear; i >= (pYear - 20); i--)
                {
                    ListaAnio.Add(i, Convert.ToString(i));
                }

                cboAnio.DisplayMember = "Value";
                cboAnio.ValueMember = "Key";
                cboAnio.DataSource = ListaAnio.ToArray();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        static IReadOnlyCollection<IWebElement> FindElements(By by)
        {
            Stopwatch w = Stopwatch.StartNew();

            while (w.ElapsedMilliseconds < 10 * 1000)
            {
                var elements = driver.FindElements(by);

                if (elements.Count > 0)
                    return elements;

                Thread.Sleep(10);
            }

            return new ReadOnlyCollection<IWebElement>(new List<IWebElement>());
        }

        static IReadOnlyCollection<IWebElement> FindDateElements(By by, ClsETCeSunat pTCeSunat)
        {
            Console.WriteLine("-----------------------------------------");
            Stopwatch w = Stopwatch.StartNew();

            while (w.ElapsedMilliseconds < 20 * 1000)
            {
                Console.WriteLine(" -> "+ w.ElapsedMilliseconds.ToString());

                var elements = driver.FindElements(by);

                if (elements.Count > 0)
                {
                    int pTCSanio = Convert.ToInt32(driver.FindElement(By.XPath("//*[@class='js-cal-years btn btn-link disabled']")).Text);
                    string pTCSnombreMes = driver.FindElement(By.XPath("//*[@class='js-cal-option btn btn-link disabled']")).Text;
                    Console.WriteLine("" + pTCSanio.ToString() + "|"+ pTCSnombreMes);

                    int pTCSmes = pTCSnombreMes == "Enero" ? 1 : pTCSnombreMes == "Febrero" ? 2 : pTCSnombreMes == "Marzo" ? 3 :
                        pTCSnombreMes == "Abril" ? 4 : pTCSnombreMes == "Mayo" ? 5 : pTCSnombreMes == "Junio" ? 6 :
                        pTCSnombreMes == "Julio" ? 7 : pTCSnombreMes == "Agosto" ? 8 : pTCSnombreMes == "Septiembre" ? 9 :
                        pTCSnombreMes == "Octubre" ? 10 : pTCSnombreMes == "Noviembre" ? 11 : pTCSnombreMes == "Diciembre" ? 12 : 0;

                    if (pTCeSunat.TCSanio == pTCSanio & pTCeSunat.TCSmes == pTCSmes)
                        return elements;

                }
                Console.WriteLine(" + 10 ");
                Thread.Sleep(10);
                
            }

            return new ReadOnlyCollection<IWebElement>(new List<IWebElement>());
        }

        private void btnConsultar_Click(object sender, EventArgs e)
        {
            try
            {
                ClsETCeSunat iTCeSunat = new ClsETCeSunat();

                driver.Navigate().GoToUrl("https://e-consulta.sunat.gob.pe/cl-at-ittipcam/tcS01Alias");

                iTCeSunat.TCSanio = Convert.ToInt32(FindElements(By.XPath("//*[@class='js-cal-years btn btn-link disabled']")).FirstOrDefault().Text);

                driver.FindElement(By.XPath("//*[@id=\"fecAsistenciaBusqDiv\"]/span")).Click();

                if (iTCeSunat.TCSanio != pTCeSunat.TCSanio)
                {
                    int pDif = iTCeSunat.TCSanio - pTCeSunat.TCSanio;

                    for (int i = 0; i < pDif; i++)
                        driver.FindElement(By.XPath("(.//*[@class='prev'])[2]")).Click();
                }

                string pMes = pTCeSunat.TCSmes == 1 ? "ene" : pTCeSunat.TCSmes == 2 ? "feb" : pTCeSunat.TCSmes == 3 ? "mar" : pTCeSunat.TCSmes == 4 ? "abr" : pTCeSunat.TCSmes == 5 ? "may" : pTCeSunat.TCSmes == 6 ? "jun" : pTCeSunat.TCSmes == 7 ? "jul" : pTCeSunat.TCSmes == 8 ? "ago" : pTCeSunat.TCSmes == 9 ? "sep" : pTCeSunat.TCSmes == 10 ? "oct" : pTCeSunat.TCSmes == 11 ? "nov" : pTCeSunat.TCSmes == 12 ? "dic" : "";

                driver.FindElement(By.XPath("//span[contains(text(),'" + pMes + ".')]")).Click();

                driver.FindElement(By.XPath("//*[@id=\"btnBuscarAsistencias\"]")).Click();



                //Nueva Ventana
                var pAnioMes = FindDateElements(By.XPath("//*[@class='js-cal-years btn btn-link disabled']"), pTCeSunat);
                //Console.WriteLine(pAnio.FirstOrDefault().Text);

                iTCeSunat.TCSanio = Convert.ToInt32(driver.FindElement(By.XPath("//*[@class='js-cal-years btn btn-link disabled']")).Text);
                iTCeSunat.TCSnombreMes = driver.FindElement(By.XPath("//*[@class='js-cal-option btn btn-link disabled']")).Text;

                iTCeSunat.TCSmes = iTCeSunat.TCSnombreMes == "Enero" ? 1 : iTCeSunat.TCSnombreMes == "Febrero" ? 2 : iTCeSunat.TCSnombreMes == "Marzo" ? 3 :
                    iTCeSunat.TCSnombreMes == "Abril" ? 4 : iTCeSunat.TCSnombreMes == "Mayo" ? 5 : iTCeSunat.TCSnombreMes == "Junio" ? 6 :
                    iTCeSunat.TCSnombreMes == "Julio" ? 7 : iTCeSunat.TCSnombreMes == "Agosto" ? 8 : iTCeSunat.TCSnombreMes == "Septiembre" ? 9 :
                    iTCeSunat.TCSnombreMes == "Octubre" ? 10 : iTCeSunat.TCSnombreMes == "Noviembre" ? 11 : iTCeSunat.TCSnombreMes == "Diciembre" ? 12 : 0;


                List<ClsETCeSunat> pTCeSunats = _LeerDatos(iTCeSunat);

                if (pTCeSunats.Where(TCS => TCS.TCSdia == 1 & TCS.TCScompra == "0.00" & TCS.TCSventa == "0.00").Count() > 0)
                {
                    ClsETCeSunat tTCeSunat = new ClsETCeSunat();

                    if (iTCeSunat.TCSmes == 1)
                    {
                        tTCeSunat.TCSanio = iTCeSunat.TCSanio - 1;
                        tTCeSunat.TCSmes = 1;
                    }
                    else
                    {
                        tTCeSunat.TCSanio = iTCeSunat.TCSanio;
                        tTCeSunat.TCSmes = iTCeSunat.TCSmes - 1;
                    }

                    driver.FindElement(By.XPath("//*[@id=\"fecAsistenciaBusqDiv\"]/span")).Click();

                    if (iTCeSunat.TCSanio != tTCeSunat.TCSanio)
                        driver.FindElement(By.XPath("(.//*[@class='prev'])[2]")).Click();

                    string tMes = tTCeSunat.TCSmes == 1 ? "ene" : tTCeSunat.TCSmes == 2 ? "feb" : tTCeSunat.TCSmes == 3 ? "mar" : tTCeSunat.TCSmes == 4 ? "abr" : tTCeSunat.TCSmes == 5 ? "may" : tTCeSunat.TCSmes == 6 ? "jun" : tTCeSunat.TCSmes == 7 ? "jul" : tTCeSunat.TCSmes == 8 ? "ago" : tTCeSunat.TCSmes == 9 ? "sep" : tTCeSunat.TCSmes == 10 ? "oct" : tTCeSunat.TCSmes == 11 ? "nov" : tTCeSunat.TCSmes == 12 ? "dic" : "";

                    driver.FindElement(By.XPath("//span[contains(text(),'" + tMes + ".')]")).Click();

                    driver.FindElement(By.XPath("//*[@id=\"btnBuscarAsistencias\"]")).Click();

                    var tAnioMes = FindDateElements(By.XPath("//*[@class='js-cal-years btn btn-link disabled']"), tTCeSunat);

                    List<ClsETCeSunat> tTCeSunats = _LeerDatos(tTCeSunat);
                    ClsETCeSunat aETCeSunat = tTCeSunats.FirstOrDefault(TCS => TCS.TCSdia == tTCeSunats.Max(mTCS => mTCS.TCSdia));

                    foreach (ClsETCeSunat pETCeSunat in pTCeSunats.Where(TCS => TCS.TCScompra == "0.00" & TCS.TCSventa == "0.00").ToList())
                    {
                        pETCeSunat.TCScompra = aETCeSunat.TCScompra;
                        pETCeSunat.TCSventa = aETCeSunat.TCSventa;
                    }
                }

                dataGridView1.DataSource = pTCeSunats;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al Guardar \n" + ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<ClsETCeSunat> _LeerDatos(ClsETCeSunat iTCeSunat)
        {
            List<ClsETCeSunat> TCeSunats = new List<ClsETCeSunat>();

            ClsETCeSunat tClsETCeSunat = new ClsETCeSunat();

            for (int i = 1; i <= DateTime.DaysInMonth(iTCeSunat.TCSanio, iTCeSunat.TCSmes); i++)
            {
                ClsETCeSunat pClsETCeSunat = new ClsETCeSunat();

                pClsETCeSunat.TCSfecha = DateTime.Parse(i + "/" + iTCeSunat.TCSmes + "/" + iTCeSunat.TCSanio, new CultureInfo("es-ES"), DateTimeStyles.NoCurrentDateDefault);

                string pFindeSemana = pClsETCeSunat.TCSfecha.ToString("dddd", new CultureInfo("es-ES")) == "domingo" ? "c-sunday" : pClsETCeSunat.TCSfecha.ToString("dddd", new CultureInfo("es-ES")) == "sábado" ? "c-saturday" : "";

                string pSelected = "";
                string pHoy = "";
                if (DateTime.Today.Month == iTCeSunat.TCSmes)
                {
                    pSelected = DateTime.Today == pClsETCeSunat.TCSfecha ? "selected" : i == 1 ? "selected" : "";
                    pHoy = DateTime.Today == pClsETCeSunat.TCSfecha ? " today" : "";
                }
                else if (i == 1)
                    pSelected = "selected";

                var pDiaCalendario = driver.FindElement(By.XPath("//*[@class='table-bordered calendar-day current _" + iTCeSunat.TCSanio + "_" + iTCeSunat.TCSmes + "_" + i + " " + pSelected + " " + pFindeSemana + " js-cal-option" + pHoy + "']"));

                pClsETCeSunat.TCSanio = iTCeSunat.TCSanio;
                pClsETCeSunat.TCSmes = iTCeSunat.TCSmes;
                pClsETCeSunat.TCSdia = i;

                string pValores = pDiaCalendario.Text;
                pClsETCeSunat.TCScompra = tClsETCeSunat.TCScompra;
                pClsETCeSunat.TCSventa = tClsETCeSunat.TCSventa;
                if (pValores.Length > 5)
                {
                    char[] delimiterChars = { '\n', '\r' };

                    string[] words = pValores.Split(delimiterChars);

                    tClsETCeSunat.TCScompra = words[2].Replace("Compra ", "");
                    tClsETCeSunat.TCSventa = words[4].Replace("Venta ", "");

                    pClsETCeSunat.TCScompra = tClsETCeSunat.TCScompra;
                    pClsETCeSunat.TCSventa = tClsETCeSunat.TCSventa;
                }

                TCeSunats.Add(pClsETCeSunat);
            }

            return TCeSunats;
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            driver.Close();
            this.Close();
        }

        private void cboMes_SelectionChangeCommitted(object sender, EventArgs e)
        {
            pTCeSunat.TCSmes = Convert.ToInt32(cboMes.SelectedValue.ToString());
            
        }

        private void cboAnio_SelectionChangeCommitted(object sender, EventArgs e)
        {
            pTCeSunat.TCSanio = Convert.ToInt32(cboAnio.SelectedValue.ToString());
        }
    }
}
