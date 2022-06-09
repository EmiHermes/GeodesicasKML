using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
//using System.Xml.Linq;
using System.Xml.XPath;
using System.IO;

namespace GeodesicasKML
{
    /// <summary>
    /// Clase para manejar el formulario.
    /// </summary>
    public partial class Form1 : Form
    {
        #region Variables
        /// <summary>
        /// Ruta del fichero que vamos a procesar.
        /// </summary>
        public string Fichero = "";
        /// <summary>
        /// XML que vamos a leer.
        /// </summary>
        XmlDocument xml = new XmlDocument();
        /// <summary>
        /// Fichero de destino.
        /// </summary>
        string ficheroDestino = @"C:\Ficheros\Salida.csv";
        #endregion
        
        string ScreenOverlayName = string.Empty;
        string ScreenOverlayIcon = string.Empty;
        string PointStyle = string.Empty;
        string PointStyle1 = string.Empty;
        string PointStyle2 = string.Empty;
        string PointStyle3 = string.Empty;
        string PointStyle4 = string.Empty;
        string PointStylePhoto = string.Empty;

        private double scale1X = 0;
        private double scale1Y = 0;
        private double scaleIcon = 0;
        private double scaleLabel = 0;

        private bool bAbierto = false;
        private int intAbierto = 590;
        private int intCerrado = 260;


        /// <summary>
        /// Carga del formulario.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            // Formulario cerrado
            this.Width = intCerrado;

            // Carga de valores de estilo
            PointStyle1 = Properties.Resources.rsPointStyle1;
            PointStyle2 = Properties.Resources.rsPointStyle2;
            PointStyle3 = Properties.Resources.rsPointStyle3;
            PointStyle4 = Properties.Resources.rsPointStyle4;
            PointStylePhoto = Properties.Resources.rsPointStylePhoto;
            ScreenOverlayName = Properties.Resources.rsScreenOverlayName;
            ScreenOverlayIcon = Properties.Resources.rsScreenOverlayIcon;

            // Carga de valores de las escalas
            // Escala del Logotipo
            scale1X = double.Parse(Properties.Resources.rsLogoScaleX.Replace(".", ","));
            scale1Y = double.Parse(Properties.Resources.rsLogoScaleY.Replace(".", ","));
            // Escala del punto y su etiqueta
            scaleIcon = double.Parse(Properties.Resources.rsPointStyleIconScale.Replace(".", ","));
            scaleLabel = double.Parse(Properties.Resources.rsPointStyleLabelScale.Replace(".", ","));

            Environment.Version.ToString();

            // Valores en las cajas de texto
            cmbPointStyle.Items.Clear();
            cmbPointStyle.Items.Add(PointStyle1);
            cmbPointStyle.Items.Add(PointStyle2);
            cmbPointStyle.Items.Add(PointStyle3);
            cmbPointStyle.Items.Add(PointStyle4);
            cmbPointStyle.SelectedIndex = 0;

            txtPointStylePhoto.Text = PointStylePhoto;
            txtName.Text = ScreenOverlayName;
            txtIcon.Text = ScreenOverlayIcon;

            #region Icons loaded from website
            try
            {
                pbImagen1.Load(PointStyle1);
            }
            catch (Exception)
            {
                pbImagen1.Image = Properties.Resources.Error;
            }
            try
            {
                pbImagen2.Load(PointStyle2);
            }
            catch (Exception)
            {
                pbImagen2.Image = Properties.Resources.Error;
            }
            try
            {
                pbImagen3.Load(PointStyle3);
            }
            catch (Exception)
            {
                pbImagen3.Image = Properties.Resources.Error;
            }
            try
            {
                pbImagen4.Load(PointStyle4);
            }
            catch (Exception)
            {
                pbImagen4.Image = Properties.Resources.Error;
            }
            #endregion

            txtScale1X.Text = scale1X.ToString("0.00");
            txtScale1Y.Text = scale1Y.ToString("0.00");
            txtIconScale.Text = scaleIcon.ToString("0.00");
            txtLabelScale.Text = scaleLabel.ToString("0.00");
        }

        #region Elementos de la ventana
        /// <summary>
        /// Constructor de la ventana.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            this.Width = intCerrado;

            // Carga de valores de estilo
            PointStyle1 = Properties.Resources.rsPointStyle1;
            PointStyle2 = Properties.Resources.rsPointStyle2;
            PointStyle3 = Properties.Resources.rsPointStyle3;
            PointStyle4 = Properties.Resources.rsPointStyle4;
            PointStylePhoto = Properties.Resources.rsPointStylePhoto;
            ScreenOverlayName = Properties.Resources.rsScreenOverlayName;
            ScreenOverlayIcon = Properties.Resources.rsScreenOverlayIcon;

            // Valores en las cajas de texto
            // Direccion del icono del usuario
            txtIconos.Text = Properties.Resources.rsIconoWebUsuario;
            // Posibles iconos
            cmbPointStyle.Items.Clear();
            cmbPointStyle.Items.Add(PointStyle1);
            cmbPointStyle.Items.Add(PointStyle2);
            cmbPointStyle.Items.Add(PointStyle3);
            cmbPointStyle.Items.Add(PointStyle4);
            cmbPointStyle.SelectedIndex = 0;

            txtPointStylePhoto.Text = PointStylePhoto;
            txtName.Text = ScreenOverlayName;
            txtIcon.Text = ScreenOverlayIcon;

            rdImagen1.Checked = true;
        }
        /// <summary>
        /// Abre/Cierra la parte para configurar los valores.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfig_Click(object sender, EventArgs e)
        {
            if (bAbierto)
            {
                ActiveForm.Width = intCerrado;

                // ScreenOverlay - Name
                ScreenOverlayName = txtName.Text.Trim();
                // ScreenOverlay - Icon
                ScreenOverlayIcon = txtIcon.Text.Trim();
                // pointStyle
                //PointStyle = cmbPointStyle.SelectedItem.ToString().Trim();
                if (rdImagen1.Checked) { PointStyle = PointStyle1; }
                if (rdImagen2.Checked) { PointStyle = PointStyle2; }
                if (rdImagen3.Checked) { PointStyle = PointStyle3; }
                if (rdImagen4.Checked) { PointStyle = PointStyle4; }
                if (rdImagen0.Checked) { PointStyle = txtIconos.Text.Trim(); }


                // pointStylePhoto
                PointStylePhoto = txtPointStylePhoto.Text.Trim();
            }
            else
            {
                ActiveForm.Width = intAbierto;
            }
            bAbierto = !bAbierto;
        }
        /// <summary>
        /// Salir de la aplicación.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSalir_Click(object sender, EventArgs e)
        {
            // Fin del programa sin errores
            Application.Exit();
        }
        /// <summary>
        /// Buscamos y procesamos el fichero KML.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnKML_Click(object sender, EventArgs e)
        {
            if (!bAbierto) bAbierto = true;
            btnConfig_Click(sender, e);
            BuscaFicheroKML();
        }
        /// <summary>
        /// Buscamos y procesamos el fichero TXT.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTXT_Click(object sender, EventArgs e)
        {
            BuscaFicheroTXT();
        }
        /// <summary>
        /// Buscamos y procesamos el fichero PRN.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPRN_Click(object sender, EventArgs e)
        {
            if (!bAbierto) bAbierto = true;
            btnConfig_Click(sender, e);
            BuscaFicheroPRN();
        }
        /// <summary>
        /// When txtIconos text is changed....
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">event.</param>
        private void txtIconos_TextChanged(object sender, EventArgs e)
        {
            rdImagen0.Checked = true;
        }
        /// <summary>
        /// When txtIconos is selected....
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">event.</param>
        private void txtIconos_MouseClick(object sender, MouseEventArgs e)
        {
            rdImagen0.Checked = true;
        }
        /// <summary>
        /// What to do if any key is pressed.
        /// </summary>
        /// <param name="msg">Message.</param>
        /// <param name="keyData">Key pressed.</param>
        /// <returns></returns>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            switch (keyData)
            {
                case Keys.F1:
                    frmAbout about = new frmAbout();
                    about.ShowDialog();
                    break;
                default:
                    break;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion

        #region Metodos Privados
        #region KML-TXT
        /// <summary>
        /// Buscamos un KML y generamos un TXT.
        /// </summary>
        private void BuscaFicheroKML()
        {
            string fichero = string.Empty;

            OpenFileDialog openFile = new OpenFileDialog();
            openFile.AddExtension = true;
            openFile.Filter = "KML files|*.kml";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Fichero = openFile.FileName;
                string fileName = openFile.SafeFileName;
                List<DatosPunto> ListaPuntos = new List<DatosPunto>();

                try
                {
                    xml.Load(Fichero);

                    XmlNode nodo = xml.SelectSingleNode("kml");
                    XmlNodeList listaLugares = xml.SelectNodes("//Placemark");

                    if (nodo != null)
                    {
                        // Recorremos los nodos encontrados
                        foreach (XmlNode Unlugar in listaLugares)
                        {
                            try
                            {
                                // Estilo del punto
                                string estilo = Unlugar.SelectSingleNode("styleUrl").InnerText;
                                // Nombre del punto
                                string name = Unlugar.SelectSingleNode("name").InnerText;
                                // Coordenadas del punto
                                string coord = Unlugar.SelectSingleNode("Point/coordinates").InnerText;
                                // Creamos el punto
                                DatosPunto puntoLocalizado = CreaPunto(estilo, name, coord);

                                // Añadimos el punto a la lista de puntos
                                ListaPuntos.Add(puntoLocalizado);
                            }
                            catch (Exception ex)
                            {
                                // Punto con datos incorrectos, nos lo saltamos
                            }
                        }

                        // Creamos un fichero de texto con lo encontrado
                        string texto = GeneraFicheroaTXT(ListaPuntos);

                        // Aviso para separar los pasos
                        MessageBox.Show(Properties.Resources.rsFileDestionationMsg, Properties.Resources.rsReadFile, MessageBoxButtons.OK, MessageBoxIcon.Question);

                        // Grabamos el fichero generado
                        bool grabado = GrabaFicheroaTXT(texto);

                        // FIN
                        if (grabado)
                        {
                            MessageBox.Show(Properties.Resources.rsCreatedFileMsg, Properties.Resources.rsCreatedFile, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show(Properties.Resources.rsNothingDoneMsg, Properties.Resources.rsNoCreatedFile, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show(Properties.Resources.rsBadFileKMLFormatMsg, Properties.Resources.rsFileError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                catch (Exception ex)
                {
                    xml = null;
                    MessageBox.Show(Properties.Resources.rsNoXMLValidMsg, Properties.Resources.rsNofile, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(Properties.Resources.rsNoFileMsg, Properties.Resources.rsNofile, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            Console.WriteLine("");
        }

        /// <summary>
        /// Creamos un punto con la información pasada.
        /// </summary>
        /// <param name="estilo">Estilo del punto.</param>
        /// <param name="name">Nombre de punto.</param>
        /// <param name="coord">Coordenadas del punto.</param>
        /// <returns>
        ///   Objeto con los datos llenos.
        /// </returns>
        private DatosPunto CreaPunto(string estilo, string name, String coord)
        {
            string[] separador = { Properties.Resources.rsSeparadorDeKML };
            // Dividimos las coordenadas en partes
            string[] coordenadas = coord.Split(separador, StringSplitOptions.RemoveEmptyEntries);

            #region Coordenada X
            // Coordenada X
            bool xNegativo = false;
            double x = Convert.ToDouble((coordenadas[0]).Replace(".", ","));
            if (x < 0)
            {
                x = -x;
                xNegativo = true;
            }
            int xGrad = (int)x;
            double xResto = x - xGrad;
            xResto = xResto * 60;
            int xMin = (int)xResto;
            xResto = xResto - xMin;
            double xSeg = xResto * 60;
            string puntoX = xGrad.ToString("00") + " " + xMin.ToString("00") + " " + xSeg.ToString("00.000000") + " ";
            puntoX = puntoX.Replace(",", ".");
            if (xNegativo)
            {
                puntoX += "W ";
            }
            else
            {
                puntoX += "E ";
            }
            #endregion

            #region Coordenada Y
            // Coordenada Y
            bool yNegativo = false;
            double y = Convert.ToDouble((coordenadas[1]).Replace(".", ","));
            if (y < 0)
            {
                y = -y;
                yNegativo = true;
            }
            int yGrad = (int)y;
            double yResto = y - yGrad;
            yResto = yResto * 60;
            int yMin = (int)yResto;
            yResto = yResto - yMin;
            double ySeg = yResto * 60;
            string puntoY = yGrad.ToString("00") + " " + yMin.ToString("00") + " " + ySeg.ToString("00.000000") + " ";
            puntoY = puntoY.Replace(",", ".");
            if (yNegativo)
            {
                puntoY += "S ";
            }
            else
            {
                puntoY += "N ";
            }
            #endregion

            #region Coordenada Z
            // Coordenada Z
            double z = Convert.ToDouble((coordenadas[2]).Replace(".", ","));
            string puntoZ = z.ToString("00.0000").Replace(",", ".");
            #endregion

            // Creamos el objeto
            DatosPunto punto = new DatosPunto();
            // Metemos los valores
            punto.Estilo = estilo.Trim();
            punto.Nombre = name.Trim();
            punto.X = Convert.ToString(puntoX.Trim());
            punto.Y = Convert.ToString(puntoY.Trim());
            punto.Z = Convert.ToString(puntoZ);

            // Devolvemos el objeto creado
            return punto;
        }

        /// <summary>
        /// Creamos una cadena de texto con todos los puntos.
        /// </summary>
        /// <param name="ListaPuntos">Lista de puntos.</param>
        /// <returns>
        ///   Cadena con todos los puntos.
        /// </returns>
        private string GeneraFicheroaTXT(List<DatosPunto> ListaPuntos)
        {
            // Cadena con todos los puntos metidos
            string response = string.Empty;
            // Cadena de separación de campos
            string separador = Properties.Resources.rsSeparadorATXT;
            // Cadena de fin de lñinea
            string FinLinea = Environment.NewLine;

            // Cabecera
            response = Properties.Resources.rsPRNStart + FinLinea;

            // Recorremos los puntos que tenemos
            foreach (DatosPunto punto in ListaPuntos)
            {
                // Creamos una línea
                //string linea = punto.Nombre + separador + punto.Estilo + separador + punto.X + separador + punto.Y + separador + punto.Z + FinLinea;
                string linea = Properties.Resources.rsPRNLineaInicio + punto.Nombre + Repite(13, " ");
                linea += punto.Y + Repite(2, " ") + punto.X + Repite(3, " ");
                linea += punto.Z + Repite(10, " ");
                linea += Properties.Resources.rsPRNLineaFin + FinLinea;

                // Añadimos la nueva línea al fichero
                response += linea;
            }

            // Devolvemos el texto generado
            return response;
        }

        /// <summary>
        /// Grabamos el texto que hemos generado.
        /// </summary>
        /// <param name="texto">Texto a grabar.</param>
        private bool GrabaFicheroaTXT(string texto)
        {
            // Elegimos el fichero de destino
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.AddExtension = true;
            saveFile.Filter = "Asc files|*.asc";
            if (saveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ficheroDestino = saveFile.FileName;

                // Grabamos el fichero
                WriteFile(ficheroDestino, texto);

                return true;
            }
            else
            {
                MessageBox.Show(Properties.Resources.rsStartAgainMsg, Properties.Resources.rsStartAgain, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }

        /// <summary>
        /// Grabamos el fichero en el ordenador.
        /// </summary>
        /// <param name="path">Ubicación y nombre del fichero.</param>
        /// <param name="str">Contenido a grabar.</param>
        public static void WriteFile(string path, string str)
        {
            // Creamos el fichero
            StreamWriter myStream = null;

            try
            {
                myStream = File.CreateText(path);
                myStream.Write(str);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (myStream != null)
                    myStream.Close();
            }
        }
        #endregion

        #region TXT-KML
        /// <summary>
        /// Buscamos un TXT y generamos un KML.
        /// </summary>
        private void BuscaFicheroTXT()
        {
            string fichero = string.Empty;

            OpenFileDialog openFile = new OpenFileDialog();
            openFile.AddExtension = true;
            openFile.Filter = "TXT files|*.txt";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Fichero = openFile.FileName;
                string fileName = openFile.FileName;
                List<DatosPunto> ListaPuntos = new List<DatosPunto>();

                if (fileName != null)
                {
                    if (fileName.EndsWith(".txt"))
                    {
                        try
                        {
                            string estilo = Properties.Resources.rsPointStyleName;

                            System.IO.StreamReader sr = new System.IO.StreamReader(fileName);
                            string line = string.Empty;
                            while ((line = sr.ReadLine()) != null)
                            {
                                DatosPunto punto = LeePuntoDeTXT(estilo, line);
                                if (punto != null)
                                {
                                    ListaPuntos.Add(punto);
                                }
                            }
                            sr.Close();

                            // Ahora creamos el XML de destino con los datos obtenidos
                            XmlDocument xmlKml = GeneraFicheroaKML(ListaPuntos);

                            if (xmlKml != null)
                            {
                                // Aviso para separar los pasos
                                MessageBox.Show(Properties.Resources.rsFileDestionationMsg, Properties.Resources.rsReadFile, MessageBoxButtons.OK, MessageBoxIcon.Question);

                                // Graba el XML generado
                                bool grabado = GrabaFicheroaKML(xmlKml);

                                // FIN
                                if (grabado)
                                {
                                    MessageBox.Show(Properties.Resources.rsCreatedFileMsg, Properties.Resources.rsCreatedFile, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    MessageBox.Show(Properties.Resources.rsNothingDoneMsg, Properties.Resources.rsNoCreatedFile, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, Properties.Resources.rsFileError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show(Properties.Resources.rsBadFileTXTFormatMsg, Properties.Resources.rsFileError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show(Properties.Resources.rsNoFileMsg, Properties.Resources.rsNofile, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show(Properties.Resources.rsNoFileMsg, Properties.Resources.rsNofile, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            Console.WriteLine("");
        }

        /// <summary>
        /// Creamos un punto con la información pasada de un TXT.
        /// </summary>
        /// <param name="estilo">Estilo del punto.</param>
        /// <param name="linea">Datos del punto, nombre y coordenadas.</param>
        /// <returns>
        ///   Objeto con los datos llenos.
        /// </returns>
        private DatosPunto LeePuntoDeTXT(string estilo, string linea)
        {
            try
            {
                string[] separador = { Properties.Resources.rsSeparadorATXT };
                // Dividimos las coordenadas en partes
                string[] coordenadas = linea.Split(separador, StringSplitOptions.RemoveEmptyEntries);

                // Creamos el objeto
                DatosPunto punto = new DatosPunto();
                // Metemos los valores
                punto.Estilo = estilo;
                punto.Nombre = Convert.ToString(coordenadas[0]);
                punto.X = Convert.ToString(coordenadas[2]);
                punto.Y = Convert.ToString(coordenadas[3]);
                punto.Z = Convert.ToString(coordenadas[4]);

                return punto;
            }
            catch (Exception ex)
            {
                // Punto malo
                return null;
            }
        }

        /// <summary>
        /// Genera el XML con la información pasada.
        /// </summary>
        /// <param name="listaPuntos">Lista de puntos a convertir en nodos.</param>
        /// <returns>
        ///   XML con los datos de los puntos añadidos. 
        /// </returns>
        private XmlDocument GeneraFicheroaKML(List<DatosPunto> listaPuntos)
        {
            try
            {
                XmlDocument xmlGenerar = new XmlDocument();
                xmlGenerar.LoadXml(Properties.Resources.rsBaseKML);
                XmlNode node = null;
                string strNodoBase = Properties.Resources.rsPlacemark;

                #region Escalas configurables
                // Escala del Logo
                node = xmlGenerar.SelectSingleNode(Properties.Resources.rsLogoScalePath);
                // Escala X
                var nameAttribute = node.Attributes["x"];
                if (nameAttribute != null)
                    nameAttribute.Value = scale1X.ToString().Replace(",", ".");
                // Escala Y
                nameAttribute = node.Attributes["y"];
                if (nameAttribute != null)
                    nameAttribute.Value = scale1Y.ToString().Replace(",", ".");

                // Escala del icono del punto
                node = xmlGenerar.SelectSingleNode(Properties.Resources.rsPointStyleIconScalePath);
                if (node != null)
                {
                    node.InnerText = txtIconScale.Text.Replace(",", ".");
                }
                // Escala de la etiqueta del punto
                node = xmlGenerar.SelectSingleNode(Properties.Resources.rsPointStyleLabelScalePath);
                if (node != null)
                {
                    node.InnerText = txtLabelScale.Text.Replace(",", ".");
                }
                #endregion

                #region Paths configurables
                // ScreenOverlayName
                node = xmlGenerar.SelectSingleNode(Properties.Resources.rsPathScreenOverlayName);
                if (node != null)
                {
                    node.InnerText = ScreenOverlayName;
                }

                // ScreenOverlayIcon
                node = xmlGenerar.SelectSingleNode(Properties.Resources.rsPathScreenOverlayIcon);
                if (node != null)
                {
                    node.InnerText = ScreenOverlayIcon;
                }

                // PointStyle
                node = xmlGenerar.SelectSingleNode(Properties.Resources.rsPathPointStyle);
                if (node != null)
                {
                    node.InnerText = PointStyle;
                }

                // PointStylePhoto
                node = xmlGenerar.SelectSingleNode(Properties.Resources.rsPathPointStylePhoto);
                if (node != null)
                {
                    node.InnerText = PointStylePhoto;
                }
                #endregion

                #region Ahora las coordenadas de los puntos obtenidos
                foreach (DatosPunto datosPunto in listaPuntos)
                {
                    XmlElement NuevaTransaccion = xmlGenerar.CreateElement(Properties.Resources.rsXMLPlacemark);
                    NuevaTransaccion.InnerXml = strNodoBase;

                    // Name
                    node = NuevaTransaccion.SelectSingleNode(Properties.Resources.rsXMLPathName);
                    if (node != null)
                    {
                        node.InnerText = datosPunto.Nombre;
                    }

                    // coordinates
                    node = NuevaTransaccion.SelectSingleNode(Properties.Resources.rsXMLPathCoordinates);
                    if (node != null)
                    {
                        string strCoordenadas = datosPunto.X + Properties.Resources.rsSeparadorAKML;
                        strCoordenadas += datosPunto.Y + Properties.Resources.rsSeparadorAKML;
                        strCoordenadas += datosPunto.Z;

                        node.InnerText = strCoordenadas;
                    }

                    XmlNode node2 = xmlGenerar.SelectSingleNode(Properties.Resources.rsXMLPathFolder);
                    node2.AppendChild(NuevaTransaccion);
                }
                #endregion

                #region Y el nombre final
                XmlNode nodeName = xmlGenerar.SelectSingleNode(Properties.Resources.rsXMLPathFolder);
                XmlElement NuevoNombre = xmlGenerar.CreateElement(Properties.Resources.rsName);
                string nombre = Inputbox.Show(Properties.Resources.rsNameTittle, Properties.Resources.rsNameTittleMsg, FormStartPosition.CenterScreen);
                NuevoNombre.InnerText = nombre;
                nodeName.AppendChild(NuevoNombre);
                #endregion

                return xmlGenerar;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Properties.Resources.rsNoCreatedFile, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Grabamos el texto que hemos generado en un XML.
        /// </summary>
        /// <param name="xmlObtenido">XML a grabar.</param>
        private bool GrabaFicheroaKML(XmlDocument xmlObtenido)
        {
            // Elegimos el fichero de destino
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.AddExtension = true;
            saveFile.Filter = "Kml files|*.kml";
            if (saveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ficheroDestino = saveFile.FileName;

                // Grabamos el fichero
                xmlObtenido.Save(ficheroDestino);

                return true;
            }
            else
            {
                MessageBox.Show(Properties.Resources.rsStartAgainMsg, Properties.Resources.rsStartAgain, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }
        #endregion

        #region PRN-KML
        /// <summary>
        /// Buscamos un PRNy generamos un KML.
        /// </summary>
        private void BuscaFicheroPRN()
        {
            string fichero = string.Empty;

            OpenFileDialog openFile = new OpenFileDialog();
            openFile.AddExtension = true;
            openFile.Filter = "PRN files|*.prn";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Fichero = openFile.FileName;
                string fileName = openFile.FileName;
                List<DatosPunto> ListaPuntos = new List<DatosPunto>();

                if (fileName != null)
                {
                    if (fileName.EndsWith(".prn"))
                    {
                        try
                        {
                            string estilo = Properties.Resources.rsPointStyleName;

                            System.IO.StreamReader sr = new System.IO.StreamReader(fileName);
                            string line = string.Empty;
                            while ((line = sr.ReadLine()) != null)
                            {
                                DatosPunto punto = LeePuntoDePRN(estilo, line);
                                if (punto != null)
                                {
                                    ListaPuntos.Add(punto);
                                }
                            }
                            sr.Close();

                            // Ahora creamos el XML de destino con los datos obtenidos
                            XmlDocument xmlKml = GeneraFicheroaKML(ListaPuntos);

                            if (xmlKml != null)
                            {
                                // Aviso para separar los pasos
                                MessageBox.Show(Properties.Resources.rsFileDestionationMsg, Properties.Resources.rsReadFile, MessageBoxButtons.OK, MessageBoxIcon.Question);

                                // Graba el XML generado
                                bool grabado = GrabaFicheroaKML(xmlKml);

                                // FIN
                                if (grabado)
                                {
                                    MessageBox.Show(Properties.Resources.rsCreatedFileMsg, Properties.Resources.rsCreatedFile, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    MessageBox.Show(Properties.Resources.rsNothingDoneMsg, Properties.Resources.rsNoCreatedFile, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, Properties.Resources.rsFileError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show(Properties.Resources.rsNoTXTValidMsg, Properties.Resources.rsFileError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show(Properties.Resources.rsNoFileMsg, Properties.Resources.rsNofile, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show(Properties.Resources.rsNoFileMsg, Properties.Resources.rsNofile, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            Console.WriteLine("");
        }

        /// <summary>
        /// Creamos un punto con la información pasada de un PRN.
        /// </summary>
        /// <param name="estilo">Estilo del punto.</param>
        /// <param name="linea">Datos del punto, nombre y coordenadas.</param>
        /// <returns>
        ///   Objeto con los datos llenos.
        /// </returns>
        private DatosPunto LeePuntoDePRN(string estilo, string linea)
        {
            try
            {
                // Primera separacion (nombre y 3 coordenadas)
                string[] separador1 = { Properties.Resources.rsSeparadorAPRN1 };
                // Dividimos las coordenadas en partes
                string[] partes = linea.Split(separador1, StringSplitOptions.RemoveEmptyEntries);

                // Segunda separacion (cada coordenada)
                string[] separador2 = { Properties.Resources.rsSeparadorAPRN2 };
                // Dividimos las coordenadas X
                string[] CoordenadasX = partes[1].Split(separador2, StringSplitOptions.RemoveEmptyEntries);
                // Dividimos las coordenadas Y
                string[] CoordenadasY = partes[2].Split(separador2, StringSplitOptions.RemoveEmptyEntries);
                // Dividimos las coordenadas Z
                string[] CoordenadasZ = partes[3].Split(separador2, StringSplitOptions.RemoveEmptyEntries);

                // Cálculo de la coordenada X
                double x1 = Convert.ToDouble(CoordenadasX[0]);
                double x2 = Convert.ToDouble(CoordenadasX[1]);
                double x3 = Convert.ToDouble((CoordenadasX[2]).Replace(".", ","));
                double coordenadaX = x1 + (x2 / 60) + (x3 / 3600);
                if (CoordenadasX[3] == "S") coordenadaX = -coordenadaX;
                string strCoordenadaX = Convert.ToString(coordenadaX).Replace(",", ".");

                // Cálculo de la coordenada Y
                double y1 = Convert.ToDouble(CoordenadasY[0]);
                double y2 = Convert.ToDouble(CoordenadasY[1]);
                double y3 = Convert.ToDouble((CoordenadasY[2]).Replace(".", ","));
                double coordenadaY = y1 + (y2 / 60) + (y3 / 3600);
                if (CoordenadasY[3] == "W") coordenadaY = -coordenadaY;
                string strCoordenadaY = Convert.ToString(coordenadaY).Replace(",", ".");

                // Cálculo de la coordenada Z
                double coordenadaZ = Convert.ToDouble((CoordenadasZ[0]).Replace(".", ","));
                string strCoordenadaZ = Convert.ToString(coordenadaZ).Replace(",", ".");

                // Creamos el objeto
                DatosPunto punto = new DatosPunto();

                // Metemos los valores
                punto.Estilo = estilo;
                punto.Nombre = Convert.ToString(partes[0]);
                //Revisamos si viene N/S - W/E o W/E - N/S
                if (CoordenadasX[3] == "S" || CoordenadasX[3] == "N")
                {
                    punto.X = strCoordenadaY;
                    punto.Y = strCoordenadaX;
                }
                else
                {
                    punto.X = strCoordenadaX;
                    punto.Y = strCoordenadaY;
                }

                punto.Z = strCoordenadaZ;

                return punto;
            }
            catch (Exception ex)
            {
                // Punto malo
                return null;
            }
        }

        #endregion
        #endregion

        /// <summary>
        /// Repite una cadena un número de veces.
        /// </summary>
        /// <param name="veces">Vces que se repite.</param>
        /// <param name="cadena">cadena a repetir.</param>
        /// <returns>
        ///  Cadena formada con las repeticiones pedidas. 
        /// </returns>
        private string Repite(int veces, string cadena)
        {
            string response = string.Empty;

            for (int i = 0; i < veces; i++)
            {
                response += cadena;
            }

            return response;
        }
    }

    /// <summary>
    /// Clase para hacer una pregunta.
    /// </summary>
    public static class Inputbox
    {
        static Form f;
        static Label l;
        static TextBox t; // Elementos necesarios
        static Button b1;
        static Button b2;
        static string Valor;

        /// <summary>
        /// Objeto Estático que muestra un pequeño diálogo para introducir datos.
        /// </summary>
        /// <param name="title">Título del diálogo.</param>
        /// <param name="prompt">Texto de información.</param>
        /// <param name="posicion">Posición de inicio.</param>
        /// <returns>Devuelve la escrito en la caja de texto como string.</returns>
        public static string Show(string title, string prompt, FormStartPosition posicion)
        {

            f = new Form();
            f.Text = title;
            f.ShowIcon = false;
            f.Icon = null;
            f.KeyPreview = true;
            f.ShowInTaskbar = false;
            f.MaximizeBox = false;
            f.MinimizeBox = false;
            f.Width = 200;
            f.FormBorderStyle = FormBorderStyle.FixedDialog;
            f.Height = 120;
            f.StartPosition = posicion;
            f.KeyPress += new KeyPressEventHandler(f_KeyPress);

            l = new Label();
            l.AutoSize = true;
            l.Text = prompt;

            t = new TextBox();
            t.Width = 182;
            t.Top = 40;

            b1 = new Button();
            b1.Text = "Aceptar";
            b1.Click += new EventHandler(b1_Click);


            b2 = new Button();
            b2.Text = "Cancelar";
            b2.Click += new EventHandler(b2_Click);

            f.Controls.Add(l);
            f.Controls.Add(t);
            f.Controls.Add(b1);
            f.Controls.Add(b2);

            l.Top = 10;
            t.Left = 5;
            t.Top = 30;

            b1.Left = 5;
            b1.Top = 60;

            b2.Left = 112;
            b2.Top = 60;

            f.ShowDialog();
            return (Valor);
        }
        static void f_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (Convert.ToChar(e.KeyChar))
            {

                case ('\r'):
                    Acepta();
                    break; ;

                case (''):
                    Cancela();
                    break; ;
            }
        }
        static void b2_Click(object sender, EventArgs e)
        {
            Cancela();
        }
        static void b1_Click(object sender, EventArgs e)
        {
            Acepta();
        }
        private static string Val
        {
            get { return (Valor); }
            set { Valor = value; }
        }
        private static void Acepta()
        {
            Val = t.Text;
            f.Dispose();
        }
        private static void Cancela()
        {
            Val = null;
            f.Dispose();
        }
    }


    /// <summary>
    /// Clase con los datos del punto.
    /// </summary>
    public class DatosPunto
    {
        public string Nombre
        { get; set; }
        public string Estilo
        { get; set; }
        public string X
        { get; set; }
        public string Y
        { get; set; }
        public string Z
        { get; set; }
    }
}
