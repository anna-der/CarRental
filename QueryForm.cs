using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace CarRentalApp
{
    public class QueryForm : Form
    {
        private enum PType { Text, Date, Number }

        private class QParam
        {
            public string Name;
            public string Label;
            public PType Type;
            public string Default;
            public bool Optional;

            public QParam(string name, string label, PType type, string def, bool optional)
            {
                Name = name;
                Label = label;
                Type = type;
                Default = def;
                Optional = optional;
            }
        }

        private class QDef
        {
            public string Title;
            public string Sql;
            public QParam[] Params = new QParam[0];
            public bool Dynamic;

            public override string ToString()
            {
                return Title;
            }
        }

        //private readonly Button export_button = new Button(); //

        private const string COST = "DATEDIFF(DAY, r.IssueDate, r.ReturnDate) * c.RentalPrice";

        private const string RENTALS_JOIN =
            "FROM dbo.Rentals r " +
            "JOIN dbo.Cars c ON c.CarID = r.CarID " +
            "JOIN dbo.Clients cl ON cl.ClientID = r.ClientID ";

        private readonly ComboBox query_combo = new ComboBox();
        private readonly FlowLayoutPanel params_panel = new FlowLayoutPanel();
        private readonly TextBox sql_view = new TextBox();
        private readonly DataGridView grid = new DataGridView();
        private readonly Label status_label = new Label();
        private readonly Dictionary<string, TextBox> param_inputs = new Dictionary<string, TextBox>();

        public DataTable LastResult { get; private set; }
        public string LastTitle { get; private set; }

        public QueryForm()
        {
            Text = "Запити до бази даних";
            Size = new Size(1100, 720);
            StartPosition = FormStartPosition.CenterParent;

            query_combo.Dock = DockStyle.Top;
            query_combo.DropDownStyle = ComboBoxStyle.DropDownList;
            query_combo.Font = new System.Drawing.Font("Segoe UI", 10);
            query_combo.SelectedIndexChanged += (s, e) => BuildParamInputs();

            params_panel.Dock = DockStyle.Top;
            params_panel.Height = 80;
            params_panel.Padding = new Padding(8);
            params_panel.AutoScroll = true;

            Button run_button = new Button();
            run_button.Text = "Виконати запит";
            run_button.Dock = DockStyle.Top;
            run_button.Height = 38;
            run_button.Click += (s, e) => RunSelected();

            Button export_button = new Button();
            export_button.Text = "Експорт у PDF";
            export_button.Dock = DockStyle.Top;
            export_button.Height = 38;
            export_button.Click += (s, e) => ExportToPdf();

            sql_view.Dock = DockStyle.Top;
            sql_view.Height = 130;
            sql_view.Multiline = true;
            sql_view.ReadOnly = true;
            sql_view.ScrollBars = ScrollBars.Vertical;
            sql_view.Font = new System.Drawing.Font("Consolas", 9.5f);

            grid.Dock = DockStyle.Fill;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            status_label.Dock = DockStyle.Bottom;
            status_label.Height = 28;
            status_label.Padding = new Padding(8, 6, 8, 0);

            Controls.Add(grid);
            Controls.Add(export_button);
            Controls.Add(status_label);
            Controls.Add(sql_view);
            Controls.Add(run_button);
            Controls.Add(params_panel);
            Controls.Add(query_combo);

            foreach (QDef def in BuildQueries())
            {
                query_combo.Items.Add(def);
            }

            query_combo.SelectedIndex = 0;
        }

        private static QParam P(string name, string label, PType type, string def, bool optional = false)
        {
            return new QParam(name, label, type, def, optional);
        }

        private static List<QDef> BuildQueries()
        {
            string today = DateTime.Today.ToString("yyyy-MM-dd");
            List<QDef> list = new List<QDef>();


            list.Add(new QDef
            {
                Title = "1. Усі оренди: клієнт, авто, кількість днів, вартість (JOIN 3 таблиць)",
                Sql =
                    "SELECT r.RentalID, cl.LastName + N' ' + cl.FirstName AS Client,\r\n" +
                    "       c.Brand, c.CarType, r.IssueDate, r.ReturnDate,\r\n" +
                    "       DATEDIFF(DAY, r.IssueDate, r.ReturnDate) AS Days,\r\n" +
                    "       " + COST + " AS Cost\r\n" +
                    RENTALS_JOIN + "\r\n" +
                    "ORDER BY r.IssueDate"
            });

            list.Add(new QDef
            {
                Title = "2. Кількість оренд і виручка за марками авто (GROUP BY, COUNT, SUM)",
                Sql =
                    "SELECT c.Brand, COUNT(r.RentalID) AS RentalsCount,\r\n" +
                    "       ISNULL(SUM(" + COST + "), 0) AS Revenue\r\n" +
                    "FROM dbo.Cars c\r\n" +
                    "LEFT JOIN dbo.Rentals r ON r.CarID = c.CarID\r\n" +
                    "GROUP BY c.Brand\r\n" +
                    "ORDER BY Revenue DESC"
            });

            list.Add(new QDef
            {
                Title = "3. Клієнти: кількість оренд, сума витрат, середня тривалість (GROUP BY)",
                Sql =
                    "SELECT cl.ClientID, cl.LastName + N' ' + cl.FirstName AS Client,\r\n" +
                    "       COUNT(*) AS RentalsCount,\r\n" +
                    "       SUM(" + COST + ") AS TotalSpent,\r\n" +
                    "       AVG(DATEDIFF(DAY, r.IssueDate, r.ReturnDate) * 1.0) AS AvgDays\r\n" +
                    RENTALS_JOIN + "\r\n" +
                    "GROUP BY cl.ClientID, cl.LastName, cl.FirstName\r\n" +
                    "ORDER BY TotalSpent DESC"
            });

            list.Add(new QDef
            {
                Title = "4. Статистика за типами кузова (COUNT, AVG, SUM, MAX)",
                Sql =
                    "SELECT c.CarType, COUNT(*) AS RentalsCount,\r\n" +
                    "       AVG(c.RentalPrice) AS AvgDailyPrice,\r\n" +
                    "       MAX(c.RentalPrice) AS MaxDailyPrice,\r\n" +
                    "       SUM(" + COST + ") AS Revenue\r\n" +
                    RENTALS_JOIN + "\r\n" +
                    "GROUP BY c.CarType\r\n" +
                    "ORDER BY Revenue DESC"
            });

            list.Add(new QDef
            {
                Title = "5. Виручка за місяцями (GROUP BY YEAR, MONTH)",
                Sql =
                    "SELECT YEAR(r.IssueDate) AS [Year], MONTH(r.IssueDate) AS [Month],\r\n" +
                    "       COUNT(*) AS RentalsCount, SUM(" + COST + ") AS Revenue\r\n" +
                    RENTALS_JOIN + "\r\n" +
                    "GROUP BY YEAR(r.IssueDate), MONTH(r.IssueDate)\r\n" +
                    "ORDER BY 1, 2"
            });

            list.Add(new QDef
            {
                Title = "6. Міста клієнтів: клієнти, оренди, виручка (LEFT JOIN, GROUP BY)",
                Sql =
                    "SELECT x.City, COUNT(DISTINCT cl.ClientID) AS Clients,\r\n" +
                    "       COUNT(r.RentalID) AS RentalsCount,\r\n" +
                    "       ISNULL(SUM(" + COST + "), 0) AS Revenue\r\n" +
                    "FROM dbo.Clients cl\r\n" +
                    "CROSS APPLY (SELECT LEFT(cl.Address, CHARINDEX(N',', cl.Address + N',') - 1) AS City) x\r\n" +
                    "LEFT JOIN dbo.Rentals r ON r.ClientID = cl.ClientID\r\n" +
                    "LEFT JOIN dbo.Cars c ON c.CarID = r.CarID\r\n" +
                    "GROUP BY x.City\r\n" +
                    "ORDER BY Revenue DESC"
            });

            list.Add(new QDef
            {
                Title = "7. Постійні клієнти: більше однієї оренди (GROUP BY + HAVING)",
                Sql =
                    "SELECT cl.ClientID, cl.LastName + N' ' + cl.FirstName AS Client,\r\n" +
                    "       cl.Phone, COUNT(*) AS RentalsCount\r\n" +
                    "FROM dbo.Clients cl\r\n" +
                    "JOIN dbo.Rentals r ON r.ClientID = cl.ClientID\r\n" +
                    "GROUP BY cl.ClientID, cl.LastName, cl.FirstName, cl.Phone\r\n" +
                    "HAVING COUNT(*) > 1\r\n" +
                    "ORDER BY RentalsCount DESC"
            });


            list.Add(new QDef
            {
                Title = "8. Оренди за період (параметри: дата від / до)",
                Params = new[]
                {
                    P("dateFrom", "Видача від (рррр-мм-дд)", PType.Date, "2026-10-01"),
                    P("dateTo", "Видача до (рррр-мм-дд)", PType.Date, "2026-10-31")
                },
                Sql =
                    "SELECT r.RentalID, cl.LastName + N' ' + cl.FirstName AS Client,\r\n" +
                    "       c.Brand, r.IssueDate, r.ReturnDate, " + COST + " AS Cost\r\n" +
                    RENTALS_JOIN + "\r\n" +
                    "WHERE r.IssueDate >= @dateFrom AND r.IssueDate <= @dateTo\r\n" +
                    "ORDER BY r.IssueDate"
            });

            list.Add(new QDef
            {
                Title = "9. Авто вибраного типу з ціною вище середньої по типу (підзапит)",
                Params = new[] { P("type", "Тип кузова (Sedan / SUV / Hatchback)", PType.Text, "Sedan") },
                Sql =
                    "SELECT c.CarID, c.Brand, c.CarType, c.RentalPrice\r\n" +
                    "FROM dbo.Cars c\r\n" +
                    "WHERE c.CarType = @type\r\n" +
                    "  AND c.RentalPrice > (SELECT AVG(c2.RentalPrice)\r\n" +
                    "                       FROM dbo.Cars c2\r\n" +
                    "                       WHERE c2.CarType = @type)\r\n" +
                    "ORDER BY c.RentalPrice DESC"
            });

            list.Add(new QDef
            {
                Title = "10. Клієнти, які орендували вибрану марку (вкладений IN)",
                Params = new[] { P("brand", "Марка (початок назви)", PType.Text, "Toyota") },
                Sql =
                    "SELECT cl.ClientID, cl.LastName, cl.FirstName, cl.Address, cl.Phone\r\n" +
                    "FROM dbo.Clients cl\r\n" +
                    "WHERE cl.ClientID IN (SELECT r.ClientID\r\n" +
                    "                      FROM dbo.Rentals r\r\n" +
                    "                      JOIN dbo.Cars c ON c.CarID = r.CarID\r\n" +
                    "                      WHERE c.Brand LIKE @brand + N'%')\r\n" +
                    "ORDER BY cl.LastName"
            });

            list.Add(new QDef
            {
                Title = "11. Авто, які жодного разу не орендували (NOT EXISTS)",
                Sql =
                    "SELECT c.CarID, c.Brand, c.CarType, c.RentalPrice\r\n" +
                    "FROM dbo.Cars c\r\n" +
                    "WHERE NOT EXISTS (SELECT 1 FROM dbo.Rentals r WHERE r.CarID = c.CarID)\r\n" +
                    "ORDER BY c.CarID"
            });

            list.Add(new QDef
            {
                Title = "12. Клієнти, чиї витрати вищі за середні по клієнтах (вкладений HAVING)",
                Sql =
                    "SELECT cl.ClientID, cl.LastName + N' ' + cl.FirstName AS Client,\r\n" +
                    "       SUM(" + COST + ") AS TotalSpent\r\n" +
                    RENTALS_JOIN + "\r\n" +
                    "GROUP BY cl.ClientID, cl.LastName, cl.FirstName\r\n" +
                    "HAVING SUM(" + COST + ") >\r\n" +
                    "  (SELECT AVG(t.Total)\r\n" +
                    "   FROM (SELECT SUM(DATEDIFF(DAY, r.IssueDate, r.ReturnDate) * c.RentalPrice) AS Total\r\n" +
                    "         FROM dbo.Rentals r\r\n" +
                    "         JOIN dbo.Cars c ON c.CarID = r.CarID\r\n" +
                    "         GROUP BY r.ClientID) t)\r\n" +
                    "ORDER BY TotalSpent DESC"
            });

            list.Add(new QDef
            {
                Title = "13. Активні оренди на вибрану дату (параметр: дата)",
                Params = new[] { P("date", "Дата (рррр-мм-дд)", PType.Date, today) },
                Sql =
                    "SELECT r.RentalID, cl.LastName + N' ' + cl.FirstName AS Client, cl.Phone,\r\n" +
                    "       c.Brand, r.IssueDate, r.ReturnDate\r\n" +
                    RENTALS_JOIN + "\r\n" +
                    "WHERE @date BETWEEN r.IssueDate AND r.ReturnDate\r\n" +
                    "ORDER BY r.ReturnDate"
            });

            list.Add(new QDef
            {
                Title = "14. Динамічний пошук оренд (умови додаються лише для заповнених полів)",
                Dynamic = true,
                Params = new[]
                {
                    P("brand", "Марка", PType.Text, "", true),
                    P("type", "Тип кузова", PType.Text, "", true),
                    P("city", "Місто клієнта", PType.Text, "", true),
                    P("minPrice", "Ціна/добу від", PType.Number, "", true),
                    P("maxPrice", "Ціна/добу до", PType.Number, "", true),
                    P("dateFrom", "Видача від (рррр-мм-дд)", PType.Date, "", true)
                },
                Sql =
                    "SELECT r.RentalID, cl.LastName + N' ' + cl.FirstName AS Client, cl.Address,\r\n" +
                    "       c.Brand, c.CarType, c.RentalPrice, r.IssueDate, r.ReturnDate\r\n" +
                    RENTALS_JOIN + "\r\n" +
                    "WHERE 1 = 1"
            });

            return list;
        }

        private void BuildParamInputs()
        {
            QDef def = query_combo.SelectedItem as QDef;
            params_panel.Controls.Clear();
            param_inputs.Clear();

            if (def == null)
            {
                return;
            }

            sql_view.Text = def.Sql;

            foreach (QParam p in def.Params)
            {
                Panel field = new Panel();
                field.Width = 200;
                field.Height = 50;
                field.Margin = new Padding(4);

                Label label = new Label();
                label.Text = p.Label + (p.Optional ? " (необов'язково)" : "");
                label.Dock = DockStyle.Top;
                label.Height = 20;

                TextBox input = new TextBox();
                input.Dock = DockStyle.Top;
                input.Text = p.Default;

                param_inputs[p.Name] = input;

                field.Controls.Add(input);
                field.Controls.Add(label);
                params_panel.Controls.Add(field);
            }

            if (def.Params.Length == 0)
            {
                Label none = new Label();
                none.Text = "Цей запит не має параметрів.";
                none.AutoSize = true;
                none.Padding = new Padding(4);
                params_panel.Controls.Add(none);
            }
        }

        private Dictionary<string, object> ReadParams(QDef def)
        {
            Dictionary<string, object> values = new Dictionary<string, object>();

            foreach (QParam p in def.Params)
            {
                string text = param_inputs[p.Name].Text.Trim();

                if (text == "")
                {
                    if (p.Optional)
                    {
                        continue;
                    }

                    throw new Exception("Заповніть параметр: " + p.Label);
                }

                if (p.Type == PType.Date)
                {
                    DateTime date;

                    if (!DateTime.TryParseExact(text, new[] { "yyyy-MM-dd", "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    {
                        throw new Exception("Невірна дата в полі «" + p.Label + "». Формат: рррр-мм-дд");
                    }

                    values[p.Name] = date;
                }
                else if (p.Type == PType.Number)
                {
                    decimal number;

                    if (!decimal.TryParse(text.Replace(',', '.'), NumberStyles.Number,
                        CultureInfo.InvariantCulture, out number))
                    {
                        throw new Exception("Введіть число в полі «" + p.Label + "»");
                    }

                    values[p.Name] = number;
                }
                else
                {
                    values[p.Name] = text;
                }
            }

            if (values.ContainsKey("dateFrom") && values.ContainsKey("dateTo")
                && (DateTime)values["dateFrom"] > (DateTime)values["dateTo"])
            {
                throw new Exception("Початкова дата не може бути пізнішою за кінцеву.");
            }

            if (values.ContainsKey("minPrice") && values.ContainsKey("maxPrice")
                && (decimal)values["minPrice"] > (decimal)values["maxPrice"])
            {
                throw new Exception("Мінімальна ціна не може бути більшою за максимальну.");
            }

            return values;
        }

        private static string BuildDynamicSql(QDef def, Dictionary<string, object> values)
        {
            StringBuilder sb = new StringBuilder(def.Sql);

            if (values.ContainsKey("brand")) sb.Append("\r\n  AND c.Brand LIKE @brand + N'%'");
            if (values.ContainsKey("type")) sb.Append("\r\n  AND c.CarType = @type");
            if (values.ContainsKey("city")) sb.Append("\r\n  AND cl.Address LIKE @city + N'%'");
            if (values.ContainsKey("minPrice")) sb.Append("\r\n  AND c.RentalPrice >= @minPrice");
            if (values.ContainsKey("maxPrice")) sb.Append("\r\n  AND c.RentalPrice <= @maxPrice");
            if (values.ContainsKey("dateFrom")) sb.Append("\r\n  AND r.IssueDate >= @dateFrom");

            sb.Append("\r\nORDER BY r.IssueDate");

            return sb.ToString();
        }

        private void RunSelected()
        {
            QDef def = query_combo.SelectedItem as QDef;

            if (def == null)
            {
                return;
            }

            try
            {
                Dictionary<string, object> values = ReadParams(def);
                string sql = def.Dynamic ? BuildDynamicSql(def, values) : def.Sql;

                sql_view.Text = sql;

                using (SqlConnection connection = Database.GetConnection())
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    foreach (KeyValuePair<string, object> item in values)
                    {
                        command.Parameters.AddWithValue("@" + item.Key, item.Value);
                    }

                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable table = new DataTable();
                        adapter.Fill(table);

                        grid.DataSource = table;
                        LastResult = table;
                        LastTitle = def.Title;

                        status_label.Text = table.Rows.Count == 0
                            ? "Запит виконано, але нічого не знайдено."
                            : "Запит виконано. Кількість рядків: " + table.Rows.Count;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Помилка запиту",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void ExportToPdf()
        {
            if (LastResult == null || LastResult.Rows.Count == 0)
            {
                MessageBox.Show("Немає даних для експорту.");
                return;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "PDF files (*.pdf)|*.pdf";
               
                string fileName = LastTitle;

                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    fileName = fileName.Replace(c, '_');
                }

                dialog.FileName = fileName + ".pdf";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                try
                {
                    BaseFont baseFont = BaseFont.CreateFont(
                        @"C:\Windows\Fonts\arial.ttf",
                        BaseFont.IDENTITY_H,
                        BaseFont.EMBEDDED);

                    iTextSharp.text.Font titleFont =
                        new iTextSharp.text.Font(
                            baseFont, 16, iTextSharp.text.Font.BOLD);

                    iTextSharp.text.Font normalFont =
                        new iTextSharp.text.Font(
                            baseFont, 10, iTextSharp.text.Font.NORMAL);

                    iTextSharp.text.Document document =
                        new iTextSharp.text.Document(
                            iTextSharp.text.PageSize.A4.Rotate());

                    using (FileStream stream =
                        new FileStream(dialog.FileName, FileMode.Create))
                    {
                        PdfWriter.GetInstance(document, stream);
                        document.Open();

                        Paragraph title = new Paragraph(
                            "ЗВІТ CAR RENTAL",
                            titleFont);

                        title.Alignment = Element.ALIGN_CENTER;
                        title.SpacingAfter = 15;

                        document.Add(title);

                        document.Add(new Paragraph(
                            "Назва запиту: " + LastTitle,
                            normalFont));

                        document.Add(new Paragraph(
                            "Дата формування: " +
                            DateTime.Now.ToString("dd.MM.yyyy HH:mm"),
                            normalFont));

                        document.Add(new Paragraph("SQL-запит:", normalFont));
                        document.Add(new Paragraph(sql_view.Text, normalFont));

                        document.Add(new Paragraph(" "));

                        PdfPTable table = new PdfPTable(LastResult.Columns.Count);
                        table.WidthPercentage = 100;
                        table.HeaderRows = 1;

                        foreach (DataColumn column in LastResult.Columns)
                        {
                            PdfPCell cell = new PdfPCell(
                                new Phrase(column.ColumnName, normalFont));

                            cell.BackgroundColor = new BaseColor(220, 230, 240);
                            cell.Padding = 5;

                            table.AddCell(cell);
                        }

                        foreach (DataRow row in LastResult.Rows)
                        {
                            foreach (DataColumn column in LastResult.Columns)
                            {
                                string value = row[column] == DBNull.Value
                                    ? ""
                                    : row[column].ToString();

                                PdfPCell cell = new PdfPCell(
                                    new Phrase(value, normalFont));

                                cell.Padding = 5;
                                table.AddCell(cell);
                            }
                        }

                        document.Add(table);

                        document.Add(new Paragraph(" "));

                        document.Add(new Paragraph(
                            "Кількість записів: " + LastResult.Rows.Count,
                            normalFont));

                        document.Close();
                    }

                    MessageBox.Show(
                        "PDF-звіт успішно створено!",
                        "Успіх",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Помилка створення PDF: " + ex.Message,
                        "Помилка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

    }

}