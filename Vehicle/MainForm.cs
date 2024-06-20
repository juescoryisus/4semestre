using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace WinFormsApp
{
    public partial class MainForm : Form
    {
        private DataTable dataTable;

        public MainForm()
        {
            InitializeComponent();
            InitializeDataGridView();
        }

        private void InitializeDataGridView()
        {
            dataGridView1.AllowUserToAddRows = false;
            dataTable = new DataTable();
            dataTable.Columns.Add("Year", typeof(int));
            dataTable.Columns.Add("Quarter", typeof(int));
            dataTable.Columns.Add("Country", typeof(string));
            dataTable.Columns.Add("Sales", typeof(decimal));
            dataGridView1.DataSource = dataTable;
        }

        private void InitializeComponent()
        {
            this.btnOpenXML = new System.Windows.Forms.Button();
            this.btnOpenCSV = new System.Windows.Forms.Button();
            this.btnModif = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.txtSales = new System.Windows.Forms.TextBox();
            this.txtCountry = new System.Windows.Forms.TextBox();
            this.txtQuarter = new System.Windows.Forms.TextBox();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnClean = new System.Windows.Forms.Button();
            this.btnshow_in_sql = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOpenXML
            // 
            this.btnOpenXML.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.btnOpenXML.Location = new System.Drawing.Point(189, 365);
            this.btnOpenXML.Name = "btnOpenXML";
            this.btnOpenXML.Size = new System.Drawing.Size(113, 49);
            this.btnOpenXML.TabIndex = 29;
            this.btnOpenXML.Text = "Open XML";
            this.btnOpenXML.UseVisualStyleBackColor = false;
            this.btnOpenXML.Click += new System.EventHandler(this.btnOpenXML_Click);
            // 
            // btnOpenCSV
            // 
            this.btnOpenCSV.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.btnOpenCSV.Location = new System.Drawing.Point(48, 365);
            this.btnOpenCSV.Name = "btnOpenCSV";
            this.btnOpenCSV.Size = new System.Drawing.Size(113, 49);
            this.btnOpenCSV.TabIndex = 28;
            this.btnOpenCSV.Text = "Open CSV";
            this.btnOpenCSV.UseVisualStyleBackColor = false;
            this.btnOpenCSV.Click += new System.EventHandler(this.btnOpenCSV_Click);
            // 
            // btnModif
            // 
            this.btnModif.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.btnModif.Location = new System.Drawing.Point(48, 299);
            this.btnModif.Name = "btnModif";
            this.btnModif.Size = new System.Drawing.Size(113, 49);
            this.btnModif.TabIndex = 26;
            this.btnModif.Text = "Modif";
            this.btnModif.UseVisualStyleBackColor = false;
            this.btnModif.Click += new System.EventHandler(this.btnModif_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.btnDelete.Location = new System.Drawing.Point(189, 235);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(113, 49);
            this.btnDelete.TabIndex = 25;
            this.btnDelete.Text = "Delete";
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.btnAdd.Location = new System.Drawing.Point(48, 235);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(113, 49);
            this.btnAdd.TabIndex = 24;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click_1);
            // 
            // txtSales
            // 
            this.txtSales.Location = new System.Drawing.Point(125, 166);
            this.txtSales.Name = "txtSales";
            this.txtSales.Size = new System.Drawing.Size(142, 26);
            this.txtSales.TabIndex = 23;
            // 
            // txtCountry
            // 
            this.txtCountry.Location = new System.Drawing.Point(125, 121);
            this.txtCountry.Name = "txtCountry";
            this.txtCountry.Size = new System.Drawing.Size(142, 26);
            this.txtCountry.TabIndex = 22;
            // 
            // txtQuarter
            // 
            this.txtQuarter.Location = new System.Drawing.Point(125, 79);
            this.txtQuarter.Name = "txtQuarter";
            this.txtQuarter.Size = new System.Drawing.Size(142, 26);
            this.txtQuarter.TabIndex = 21;
            // 
            // txtYear
            // 
            this.txtYear.Location = new System.Drawing.Point(125, 39);
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(142, 26);
            this.txtYear.TabIndex = 20;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(43, 166);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 20);
            this.label4.TabIndex = 19;
            this.label4.Text = "Sales:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(43, 121);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(68, 20);
            this.label3.TabIndex = 18;
            this.label3.Text = "Country:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(44, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 20);
            this.label2.TabIndex = 17;
            this.label2.Text = "Quarter:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(43, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 20);
            this.label1.TabIndex = 16;
            this.label1.Text = "Year:";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(323, 39);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 28;
            this.dataGridView1.Size = new System.Drawing.Size(466, 440);
            this.dataGridView1.TabIndex = 15;
            // 
            // btnClean
            // 
            this.btnClean.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.btnClean.Location = new System.Drawing.Point(104, 435);
            this.btnClean.Name = "btnClean";
            this.btnClean.Size = new System.Drawing.Size(117, 44);
            this.btnClean.TabIndex = 30;
            this.btnClean.Text = "Clean";
            this.btnClean.UseVisualStyleBackColor = false;
            this.btnClean.Click += new System.EventHandler(this.btnClean_Click);
            // 
            // btnshow_in_sql
            // 
            this.btnshow_in_sql.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.btnshow_in_sql.Location = new System.Drawing.Point(189, 304);
            this.btnshow_in_sql.Name = "btnshow_in_sql";
            this.btnshow_in_sql.Size = new System.Drawing.Size(113, 44);
            this.btnshow_in_sql.TabIndex = 31;
            this.btnshow_in_sql.Text = "show in sql";
            this.btnshow_in_sql.UseVisualStyleBackColor = false;
            this.btnshow_in_sql.Click += new System.EventHandler(this.btnshow_in_sql_Click);
            // 
            // MainForm
            // 
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(818, 491);
            this.Controls.Add(this.btnshow_in_sql);
            this.Controls.Add(this.btnClean);
            this.Controls.Add(this.btnOpenXML);
            this.Controls.Add(this.btnOpenCSV);
            this.Controls.Add(this.btnModif);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.txtSales);
            this.Controls.Add(this.txtCountry);
            this.Controls.Add(this.txtQuarter);
            this.Controls.Add(this.txtYear);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "MainForm";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private Button btnOpenXML;
        private Button btnOpenCSV;
        private Button btnModif;
        private Button btnDelete;
        private Button btnAdd;
        private TextBox txtSales;
        private TextBox txtCountry;
        private TextBox txtQuarter;
        private TextBox txtYear;
        private Label label4;
        private Label label3;
        private Label label2;
        private Label label1;
        private DataGridView dataGridView1;

        private void btnAdd_Click_1(object sender, EventArgs e)
        {
            DataRow newRow = dataTable.NewRow();
            newRow["Year"] = Convert.ToInt32(txtYear.Text);
            newRow["Quarter"] = Convert.ToInt32(txtQuarter.Text);
            newRow["Country"] = txtCountry.Text;
            newRow["Sales"] = Convert.ToDecimal(txtSales.Text);
            dataTable.Rows.Add(newRow);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            // Código para eliminar fila seleccionada
            if (dataGridView1.SelectedRows.Count > 0)
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.SelectedRows[0].Index);
            }
        }

        private void btnModif_Click(object sender, EventArgs e)
        {
            // Código para modificar fila seleccionada
            if (dataGridView1.SelectedRows.Count > 0)
            {
                dataGridView1.SelectedRows[0].Cells["Year"].Value = Convert.ToInt32(txtYear.Text);
                dataGridView1.SelectedRows[0].Cells["Quarter"].Value = Convert.ToInt32(txtQuarter.Text);
                dataGridView1.SelectedRows[0].Cells["Country"].Value = txtCountry.Text;
                dataGridView1.SelectedRows[0].Cells["Sales"].Value = Convert.ToDecimal(txtSales.Text);
            }
        }

        private void btnOpenCSV_Click(object sender, EventArgs e)
        {
            // Código para abrir archivo CSV y cargar en el DataGridView
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos CSV (*.csv)|*.csv";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DataTable csvTable = new DataTable();
                    using (StreamReader sr = new StreamReader(openFileDialog.FileName))
                    {
                        string[] headers = sr.ReadLine().Split(',');
                        foreach (string header in headers)
                        {
                            csvTable.Columns.Add(header);
                        }
                        while (!sr.EndOfStream)
                        {
                            string[] rows = sr.ReadLine().Split(',');
                            DataRow dr = csvTable.NewRow();
                            for (int i = 0; i < headers.Length; i++)
                            {
                                dr[i] = rows[i];
                            }
                            csvTable.Rows.Add(dr);
                        }
                    }
                    dataGridView1.DataSource = csvTable;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al abrir el archivo CSV: " + ex.Message);
                }
            }
        }

        private void btnOpenXML_Click(object sender, EventArgs e)
        {
            // Código para abrir archivo XML y cargar en el DataGridView
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos XML (*.xml)|*.xml";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DataSet dataSet = new DataSet();
                    dataSet.ReadXml(openFileDialog.FileName);
                    dataGridView1.DataSource = dataSet.Tables[0];
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al abrir el archivo XML: " + ex.Message);
                }
            }
        }

        private Button btnClean;

        private void btnClean_Click(object sender, EventArgs e)
        {
            if (dataTable != null)
            {
                dataTable.Clear(); // Limpiar los datos pero mantener la estructura de las columnas
            }
            dataGridView1.DataSource = null;


        }

        private Button btnshow_in_sql;

        private void btnshow_in_sql_Click(object sender, EventArgs e)
        {
            // Guardar datos desde el DataGridView a SQL Server
            string connectionString = "server=LOCALHOST\\SQLEXPRESS; database=DATO; integrated security=true";
            using (SqlConnection conexion = new SqlConnection(connectionString))
            {
                try
                {
                    conexion.Open();

                    // Iterar sobre cada fila del DataGridView
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow) continue; // Saltar la fila nueva vacía al final

                        // Obtener los valores de cada celda de la fila
                        var year = row.Cells["Year"].Value;
                        var quarter = row.Cells["Quarter"].Value;
                        var county = row.Cells["Country"].Value;
                        var sales = row.Cells["Sales"].Value;

                        // Construir el comando SQL
                        string query = "INSERT INTO dato (Year, Quarter, Country, Sales) VALUES (@Year, @Quarter, @Country, @Sales)";

                        using (SqlCommand comando = new SqlCommand(query, conexion))
                        {
                            // Asignar los parámetros
                            comando.Parameters.AddWithValue("@Year", year ?? DBNull.Value);
                            comando.Parameters.AddWithValue("@Quarter", quarter ?? DBNull.Value);
                            comando.Parameters.AddWithValue("@Country", county ?? DBNull.Value);
                            comando.Parameters.AddWithValue("@Sales", sales ?? DBNull.Value);

                            // Ejecutar el comando SQL
                            comando.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Datos guardados con éxito.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar datos: " + ex.Message);
                }
                finally
                {
                    // Asegurarse de cerrar la conexión
                    if (conexion.State == ConnectionState.Open)
                    {
                        conexion.Close();
                    }
                }
            }
        }

    }
    
}
