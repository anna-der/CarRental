using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace CarRentalApp
{
    public class DataForm : Form
    {
        private readonly string table_name;
        private readonly DataGridView grid = new DataGridView();
        private readonly FlowLayoutPanel fields_panel = new FlowLayoutPanel();

        private readonly Dictionary<string, TextBox> inputs =
            new Dictionary<string, TextBox>();

        private string[] columns;
        private string[] editable_columns;

        public DataForm(string table)
        {
            table_name = table;

            switch (table_name)
            {
                case "Cars":
                    Text = "Автомобілі";
                    columns = new[] { "CarID", "Brand", "Price", "RentalPrice", "CarType" };
                    editable_columns = new[] { "Brand", "Price", "RentalPrice", "CarType" };
                    break;

                case "Clients":
                    Text = "Клієнти";
                    columns = new[] { "ClientID", "LastName", "FirstName", "MiddleName", "Address", "Phone" };
                    editable_columns = new[] { "LastName", "FirstName", "MiddleName", "Address", "Phone" };
                    break;

                case "Rentals":
                    Text = "Оренда автомобілів";
                    columns = new[] { "RentalID", "CarID", "ClientID", "IssueDate", "ReturnDate" };
                    editable_columns = new[] { "CarID", "ClientID", "IssueDate", "ReturnDate" };
                    break;

                default:
                    throw new ArgumentException("Невідома таблиця");
            }

            Size = new Size(1000, 650);
            StartPosition = FormStartPosition.CenterParent;

            grid.Dock = DockStyle.Fill;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.SelectionChanged += Grid_SelectionChanged;

            fields_panel.Dock = DockStyle.Top;
            fields_panel.Height = 115;
            fields_panel.Padding = new Padding(8);
            fields_panel.AutoScroll = true;
            fields_panel.WrapContents = true;

            foreach (string column in editable_columns)
            {
                Panel field = new Panel();
                field.Width = 180;
                field.Height = 65;
                field.Margin = new Padding(5);

                Label label = new Label();
                label.Text = column;
                label.Dock = DockStyle.Top;
                label.Height = 22;

                TextBox input = new TextBox();
                input.Dock = DockStyle.Top;
                input.Width = 165;

                inputs[column] = input;

                field.Controls.Add(input);
                field.Controls.Add(label);
                fields_panel.Controls.Add(field);
            }

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Bottom;
            buttons.Height = 55;
            buttons.Padding = new Padding(8);

            buttons.Controls.Add(CreateButton("Оновити", LoadData));
            buttons.Controls.Add(CreateButton("Додати", AddRecord));
            buttons.Controls.Add(CreateButton("Зберегти зміни", UpdateRecord));
            buttons.Controls.Add(CreateButton("Видалити", DeleteRecord));
            buttons.Controls.Add(CreateButton("Очистити поля", ClearFields));

            Controls.Add(grid);
            Controls.Add(buttons);
            Controls.Add(fields_panel);

            LoadData();


        }

        private Button CreateButton(string text, Action action)
        {
            Button button = new Button();
            button.Text = text;
            button.AutoSize = true;
            button.Height = 32;
            button.Margin = new Padding(5);
            button.Click += (sender, e) => RunSafely(action);

            return button;
        }

        private void RunSafely(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Помилка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void LoadData()
        {
            using (SqlConnection connection = Database.GetConnection())
            using (SqlDataAdapter adapter = new SqlDataAdapter(
                "SELECT * FROM dbo.[" + table_name + "]", connection))
            {
                DataTable data = new DataTable();
                adapter.Fill(data);
                grid.DataSource = data;
            }
        }

        private Dictionary<string, object> GetValues()
        {
            Dictionary<string, object> values =
                new Dictionary<string, object>();

            foreach (string column in editable_columns)
            {
                string value = inputs[column].Text.Trim();

                if (column == "MiddleName" && value == "")
                {
                    values[column] = DBNull.Value;
                }
                else if (value == "")
                {
                    throw new Exception("Заповніть поле " + column);
                }
                else if (column == "Price" || column == "RentalPrice")
                {
                    decimal number;

                    bool parsed = decimal.TryParse(
                        value,
                        System.Globalization.NumberStyles.Number,
                        System.Globalization.CultureInfo.CurrentCulture,
                        out number);

                    if (!parsed)
                    {
                        parsed = decimal.TryParse(
                            value,
                            System.Globalization.NumberStyles.Number,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out number);
                    }

                    if (!parsed)
                    {
                        throw new Exception(
                            "Введіть коректне число для поля " + column);
                    }

                    values[column] = number;
                }
                else if (column == "CarID" || column == "ClientID")
                {
                    int number;

                    if (!int.TryParse(value, out number))
                    {
                        throw new Exception(
                            "Введіть ціле число для поля " + column);
                    }

                    values[column] = number;
                }
                else if (column == "IssueDate" || column == "ReturnDate")
                {
                    DateTime date;

                    if (!DateTime.TryParse(value, out date))
                    {
                        throw new Exception(
                            "Введіть коректну дату для поля " + column);
                    }

                    values[column] = date.Date;
                }
                else
                {
                    values[column] = value;
                }
            }

            return values;
        }

        private void AddRecord()
        {
            Dictionary<string, object> values = GetValues();

            string column_list = string.Join(", ", values.Keys);

            string[] parameter_names = values.Keys
                .Select((name, position) => "@p" + position)
                .ToArray();

            string parameter_list = string.Join(", ", parameter_names);

            string sql = "INSERT INTO dbo.[" + table_name + "] (" +
                         column_list + ") VALUES (" + parameter_list + ")";

            using (SqlConnection connection = Database.GetConnection())
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                int parameter_index = 0;

                foreach (KeyValuePair<string, object> item in values)
                {
                    command.Parameters.AddWithValue(
                        "@p" + parameter_index, item.Value);

                    parameter_index++;
                }

                connection.Open();
                command.ExecuteNonQuery();
            }

            LoadData();
            ClearFields();

            MessageBox.Show("Запис успішно додано!");
        }

        private void UpdateRecord()
        {
            if (grid.CurrentRow == null)
            {
                throw new Exception("Оберіть запис у таблиці");
            }

            Dictionary<string, object> values = GetValues();
            object id = grid.CurrentRow.Cells[columns[0]].Value;

            string[] assignments = values.Keys
                .Select((name, position) => name + " = @p" + position)
                .ToArray();

            string set_clause = string.Join(", ", assignments);

            string sql = "UPDATE dbo.[" + table_name + "] SET " +
                         set_clause + " WHERE " + columns[0] + " = @id";

            using (SqlConnection connection = Database.GetConnection())
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                int parameter_index = 0;

                foreach (KeyValuePair<string, object> item in values)
                {
                    command.Parameters.AddWithValue(
                        "@p" + parameter_index, item.Value);

                    parameter_index++;
                }

                command.Parameters.AddWithValue("@id", id);

                connection.Open();

                int affected_rows = command.ExecuteNonQuery();

                if (affected_rows == 0)
                {
                    throw new Exception("Запис не знайдено або не було змінено.");
                }
            }

            LoadData();

            MessageBox.Show("Запис успішно оновлено!");
        }

        private void DeleteRecord()
        {
            if (grid.CurrentRow == null)
            {
                throw new Exception("Оберіть запис у таблиці");
            }

            DialogResult result = MessageBox.Show(
                "Ви дійсно хочете видалити вибраний запис?",
                "Підтвердження",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
            {
                return;
            }

            object id = grid.CurrentRow.Cells[columns[0]].Value;

            string sql = "DELETE FROM dbo.[" + table_name + "] WHERE " +
                         columns[0] + " = @id";

            using (SqlConnection connection = Database.GetConnection())
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                command.Parameters.AddWithValue("@id", id);

                connection.Open();
                command.ExecuteNonQuery();
            }

            LoadData();
            ClearFields();

            MessageBox.Show("Запис видалено!");
        }

        private void Grid_SelectionChanged(object sender, EventArgs e)
        {
            if (grid.CurrentRow == null || grid.CurrentRow.IsNewRow)
            {
                return;
            }

            foreach (string column in editable_columns)
            {
                if (!grid.Columns.Contains(column))
                {
                    continue;
                }

                object value = grid.CurrentRow.Cells[column].Value;

                inputs[column].Text =
                    value == null || value == DBNull.Value
                    ? ""
                    : value.ToString();
            }
        }

        private void ClearFields()
        {
            foreach (TextBox input in inputs.Values)
            {
                input.Clear();
            }

            grid.ClearSelection();
        }


    }


}


