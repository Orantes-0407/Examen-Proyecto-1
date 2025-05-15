using System.Data;
using System.Data.SqlClient;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using Microsoft.Office.Core;


namespace Proyecto1
{
    public partial class Form1 : Form
    {
        private string connStr = "Server=MELANY\\SQLEXPRESS;Database=Agente_DB;Trusted_Connection=True";
        private readonly string apiKey = "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"; // Cambia esto por tu clave real
        private readonly string endpoint = "https://api.openai.com/v1/chat/completions";

        public Form1()
        {
            InitializeComponent();
            this.txtTema = this.textBox1;
            this.rtbResultado = this.richTextBox1;
            progressBar1.Visible = false;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            string tema = txtTema.Text.Trim();
            if (string.IsNullOrEmpty(tema))
            {
                MessageBox.Show("Por favor, ingresa un tema.");
                return;
            }
            progressBar1.Visible = true;
            progressBar1.Value = 10;

            string respuesta = await ConsultarIA(tema);
            rtbResultado.Text = respuesta;
            progressBar1.Value = 40;

            GuardarEnBD(tema, respuesta);
            progressBar1.Value = 60;

            string carpeta = Path.Combine(Environment.CurrentDirectory, "Investigaciones", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            Directory.CreateDirectory(carpeta);

            string pathWord = Path.Combine(carpeta, "Informe.docx");
            GenerarWord(pathWord, tema, respuesta);
            progressBar1.Value = 80;

            string pathPpt = Path.Combine(carpeta, "Presentacion.pptx");
            GenerarPowerPoint(pathPpt, tema, respuesta);
            progressBar1.Value = 100;
            progressBar1.Visible = false;

            MessageBox.Show("Proceso completado. Archivos generados.");
        }
        private async Task<string> ConsultarIA(string tema)
        {
            var cliente = new HttpClient();
            cliente.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

            var body = new
            {
                model = "gpt-3.5-turbo",

                messages = new[]
                {
            new { role = "user", content = $"Investiga sobre el siguiente tema: {tema}" }
        }
            };

            string json = JsonConvert.SerializeObject(body);
            var contenido = new StringContent(json, Encoding.UTF8, "application/json");

            var respuesta = await cliente.PostAsync(endpoint, contenido); // Usar el campo endpoint aquí
            string jsonRespuesta = await respuesta.Content.ReadAsStringAsync();

            dynamic resultado = JsonConvert.DeserializeObject(jsonRespuesta);

            if (resultado.error != null)
            {
                MessageBox.Show("Error de la API: " + resultado.error.message);
                return "Error de la API: " + resultado.error.message;
            }

            if (resultado.choices != null &&
                resultado.choices.Count > 0 &&
                resultado.choices[0].message != null &&
                resultado.choices[0].message.content != null)
            {
                return resultado.choices[0].message.content.ToString();
            }
            else
            {
                MessageBox.Show("La respuesta de la API no tiene el formato esperado:\n" + jsonRespuesta);
                return "Error: respuesta inesperada de la API.";
            }
        }




        private void GuardarEnBD(string prompt, string respuesta)
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                string query = "INSERT INTO Investigaciones (Prompt, Respuesta, Fecha) VALUES (@p, @r, GETDATE())";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@p", prompt);
                cmd.Parameters.AddWithValue("@r", respuesta);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }
        private void GenerarWord(string ruta, string tema, string contenido)
        {
            var word = new Microsoft.Office.Interop.Word.Application();
            var doc = word.Documents.Add();

            doc.Content.Text = $"Tema de investigación: {tema}\n\n{contenido}";
            doc.SaveAs2(ruta);
            doc.Close();
            word.Quit();
        }
        private void GenerarPowerPoint(string ruta, string tema, string contenido)
        {
            var ppt = new Microsoft.Office.Interop.PowerPoint.Application();
            var pres = ppt.Presentations.Add();

            // Divide el contenido en párrafos (puedes ajustar el separador si lo necesitas)
            var partes = contenido.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < partes.Length; i++)
            {
                var slide = pres.Slides.Add(i + 1, PpSlideLayout.ppLayoutText);
                // Título: solo en la primera diapositiva, en las demás puedes poner el tema o dejarlo vacío
                slide.Shapes[1].TextFrame.TextRange.Text = (i == 0) ? tema : $"{tema} (cont.)";
                // Contenido
                slide.Shapes[2].TextFrame.TextRange.Text = partes[i].Trim();
            }

            pres.SaveAs(ruta, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoTrue);
            pres.Close();
            ppt.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(pres);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ppt);
        }







        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
