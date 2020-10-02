using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace AccessDataBaseDemo
{
    public partial class Form1 : Form
    {
        // строка подключения к MS Access
        // вариант 1
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Workers.mdb;";
        // вариант 2
        //public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Workers.mdb;";

        // поле - ссылка на экземпляр класса OleDbConnection для соединения с БД
        private OleDbConnection myConnection;

        // конструктор класса формы
        public Form1()
        {
            InitializeComponent();

            // создаем экземпляр класса OleDbConnection
            myConnection = new OleDbConnection(connectString);

            // открываем соединение с БД
            myConnection.Open();
        }

        // обработчик события закрытия формы
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // заркываем соединение с БД
            myConnection.Close();
        }

        // обработчик нажатия кнопки SELECT1
        private void selectButton1_Click(object sender, EventArgs e)
        {
            OleDbCommand command;
            OleDbDataReader reader;
            string Commercial_name = textBox1.Text;
            string Category = comboBox1.Text;
            string idCommercial;
            string Store_query;
            string Product_query;
            idCommercial = String.Format("SELECT Id_commercial FROM Commercial WHERE Commercial_name = \"{0}\"", Commercial_name);
            command = new OleDbCommand(idCommercial, myConnection);
            reader = command.ExecuteReader();
            if (!reader.HasRows)
            {
                richTextBox1.Text += "В базе не найдено такой торговой сети\n";
                return;
            }
            richTextBox1.Text += "Подключаемся к базе данных\n";
            List<string> Idshop = new List<string>();
            List<string[]> Items = new List<string[]>();
            List<string[]> Items_c = new List<string[]>();
            if (Category.Equals("Все"))
            {
                richTextBox1.Text += "Выполяем запрос\n";
                //Получаем по id торговой сети по её имени
                idCommercial = String.Format("SELECT Id_commercial FROM Commercial WHERE Commercial_name = \"{0}\"", Commercial_name);  
                command = new OleDbCommand(idCommercial, myConnection);
                reader = command.ExecuteReader();
                while (reader.Read())
                    idCommercial = reader[0].ToString();
                reader.Close();
                //Получаем список магазинов(складов) этой торговой сети
                Store_query = String.Format("SELECT Id_store FROM Shops WHERE Id_commercial = {0}", idCommercial);
                command = new OleDbCommand(Store_query, myConnection);
                reader = command.ExecuteReader();
                while (reader.Read())
                    Idshop.Add(reader[0].ToString());
                reader.Close();
                //Получаем список: id товаров, их цены и количество для каждой торговой точки(склада)
                int k = 0;
                int t;
                string categ;
                Items.Add(new string[6]);
                richTextBox1.Text += "Формируем отчёт на основе результатов выполнения запроса\n";
                // Получить объект приложения Word.
                Word._Application word_app = new Word.Application();

                // Сделать Word видимым (необязательно).
                //word_app.Visible = true;

                // Создаем документ Word.
                int ml = 0;
                object missing = Type.Missing;
                List<Word.Paragraph> para= new List<Word.Paragraph>();
                
                Word._Document word_doc = word_app.Documents.Add(
                    ref missing, ref missing, ref missing, ref missing);
                para.Add(word_doc.Paragraphs.Add(ref missing));
                // Создаем абзац заголовка.
                para[ml].Range.Text = Commercial_name;
                object style_name = "Заголовок 1";
                para[ml].Range.set_Style(ref style_name);
                para[ml].Range.InsertParagraphAfter();
                para.Add( word_doc.Paragraphs.Add(ref missing));
                ml++;
                para[ml].Range.Text = String.Format("Отчёт создан {0}",System.DateTime.Now.ToLocalTime());
                para[ml].Range.InsertParagraphAfter();
                Word.Range _currentRange = para[ml].Range;


                foreach (string i in Idshop)
                {
                    string[] info = new string[5];
                    Store_query = String.Format("SELECT Id_shop, Address, Holder_name, Holder_phone, Email FROM Shops WHERE Id_store = \"{0}\"", i);
                    command = new OleDbCommand(Store_query, myConnection);
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        info[0] = reader[0].ToString();
                        info[1] = reader[1].ToString();
                        info[2] = reader[2].ToString();
                        info[3] = reader[3].ToString();
                        info[4] = reader[4].ToString();
                    }
                    reader.Close();
                    para[ml].Range.InsertParagraphAfter();
                    
                    para[ml].Range.InsertBefore(String.Format("Торговая точка {0} \n Адрес: {1} \n Владелец: {2} \n Телефон: {3} \n Почта: {4}", info[0], info[1], info[2], info[3], info[4]));
                    _currentRange = para[ml].Range.Next();
                    k = 0;
                    Items = new List<string[]>();
                    Items.Add(new string[6]);
                    Product_query = String.Format("SELECT Id_item, Amount, Price FROM Store WHERE Id_store = {0}", i);
                    command = new OleDbCommand(Product_query, myConnection);
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        Items[k][0] = reader[0].ToString();
                        Items[k][4] = reader[1].ToString();
                        Items[k][5] = reader[2].ToString();
                        k++;
                        Items.Add(new string[6]);
                    }

                    reader.Close();
                    for (int l = 0; l < k; l++)
                    {
                        Product_query = String.Format("SELECT Item_name, Id_category,Producer_name FROM Items WHERE Id_item = {0}", Items[l][0]);
                        //richTextBox1.Text += Product_query + "  \n";
                        command = new OleDbCommand(Product_query, myConnection);
                        reader = command.ExecuteReader();
                        categ = "";
                        while (reader.Read())
                        {
                            Items[l][1] = reader[0].ToString();
                            Items[l][3] = reader[2].ToString();
                            categ = reader[1].ToString();
                        }
                        reader.Close();
                        Product_query = String.Format("SELECT Category_name FROM Item_category WHERE Id_category = {0}", categ);
                        command = new OleDbCommand(Product_query, myConnection);
                        reader = command.ExecuteReader();
                        while (reader.Read())
                            Items[l][2] = reader[0].ToString();
                        reader.Close();
                    }
                    Word.Table _table = word_doc.Tables.Add(para[ml].Range, Items.Count, 5, ref missing, ref missing);
                    _table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
                    _table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
                    _currentRange = _table.Cell(1, 1).Range;
                    _table.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorOrange;
                    _currentRange.InsertAfter("Название товара");
                    _currentRange = _table.Cell(1, 2).Range;
                    _table.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorPaleBlue;
                    _currentRange.InsertAfter("Категория");
                    _currentRange = _table.Cell(1, 3).Range;
                    _table.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 3).Shading.BackgroundPatternColor = Word.WdColor.wdColorSeaGreen;
                    _currentRange.InsertAfter("Изготовитель");
                    _currentRange = _table.Cell(1, 4).Range;
                    _table.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 4).Shading.BackgroundPatternColor = Word.WdColor.wdColorViolet;
                    _currentRange.InsertAfter("Количество товара");
                    _currentRange = _table.Cell(1, 5).Range;
                    _table.Cell(1, 5).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 5).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow;
                    _currentRange.InsertAfter("Цена");
                    for (int index = 2; index <= Items.Count; index++)
                    {
                        for (int j = 1; j <= 5; j++)
                        {
                            _currentRange = _table.Cell(index, j).Range;
                            _table.Cell(index, j).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            _currentRange.InsertAfter(Items[index - 2][j]);
                        }
                    }
                    para[ml].Range.InsertParagraphAfter();
                    para.Add(word_doc.Paragraphs.Add(ref missing));
                    ml++;
                    para[ml].Range.InsertParagraphAfter();
                    para[ml].Range.GoToEditableRange();
                    _currentRange = para[ml].Range;
                }
                
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.DefaultExt = "doc";
                sfd.Filter = "Word files (*.doc)|*.doc|All files (*.*)|*.*";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    word_doc.SaveAs2(sfd.FileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
ref missing, ref missing, ref missing, ref missing,
ref missing, ref missing, ref missing, ref missing,
 ref missing);
                }
                // Закрыть.
                object save_changes = false;
                word_doc.Close(ref save_changes, ref missing, ref missing);
                word_app.Quit(ref save_changes, ref missing, ref missing);
                richTextBox1.Text += "Готово\n";
            }
            else
            {
                richTextBox1.Text += "Выполяем запрос\n";
                //Получаем по id торговой сети по её имени
                idCommercial = String.Format("SELECT Id_commercial FROM Commercial WHERE Commercial_name = \"{0}\"", Commercial_name);
                command = new OleDbCommand(idCommercial, myConnection);
                reader = command.ExecuteReader();
                while (reader.Read())
                    idCommercial = reader[0].ToString();
                reader.Close();
                //Получаем список магазинов(складов) этой торговой сети
                Store_query = String.Format("SELECT Id_store FROM Shops WHERE Id_commercial = {0}", idCommercial);
                command = new OleDbCommand(Store_query, myConnection);
                reader = command.ExecuteReader();
                while (reader.Read())
                    Idshop.Add(reader[0].ToString());
                reader.Close();
                //Получаем список: id товаров, их цены и количество для каждой торговой точки(склада)
                int k = 0;
                int t;
                string categ;
                Items.Add(new string[6]);
                richTextBox1.Text += "Формируем отчёт на основе результатов выполнения запроса\n";
                // Получить объект приложения Word.
                Word._Application word_app = new Word.Application();

                // Сделать Word видимым (необязательно).
                //word_app.Visible = true;

                // Создаем документ Word.
                int ml = 0;
                object missing = Type.Missing;
                List<Word.Paragraph> para = new List<Word.Paragraph>();

                Word._Document word_doc = word_app.Documents.Add(
                    ref missing, ref missing, ref missing, ref missing);
                para.Add(word_doc.Paragraphs.Add(ref missing));
                // Создаем абзац заголовка.
                para[ml].Range.Text = Commercial_name;
                object style_name = "Заголовок 1";
                para[ml].Range.set_Style(ref style_name);
                para[ml].Range.InsertParagraphAfter();
                para.Add(word_doc.Paragraphs.Add(ref missing));
                ml++;
                para[ml].Range.Text = String.Format("Отчёт создан {0}", System.DateTime.Now.ToLocalTime());
                para[ml].Range.InsertParagraphAfter();
                Word.Range _currentRange = para[ml].Range;


                foreach (string i in Idshop)
                {
                    string[] info = new string[5];
                    Store_query = String.Format("SELECT Id_shop, Address, Holder_name, Holder_phone, Email FROM Shops WHERE Id_store = \"{0}\"", i);
                    command = new OleDbCommand(Store_query, myConnection);
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        info[0] = reader[0].ToString();
                        info[1] = reader[1].ToString();
                        info[2] = reader[2].ToString();
                        info[3] = reader[3].ToString();
                        info[4] = reader[4].ToString();
                    }
                    reader.Close();
                    para[ml].Range.InsertParagraphAfter();

                    para[ml].Range.InsertBefore(String.Format("Торговая точка {0} \n Адрес: {1} \n Владелец: {2} \n Телефон: {3} \n Почта: {4}", info[0], info[1], info[2], info[3], info[4]));
                    _currentRange = para[ml].Range.Next();
                    k = 0;
                    Items = new List<string[]>();
                    Items.Add(new string[6]);
                    Product_query = String.Format("SELECT Id_item, Amount, Price FROM Store WHERE Id_store = {0}", i);
                    command = new OleDbCommand(Product_query, myConnection);
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        Items[k][0] = reader[0].ToString();
                        Items[k][4] = reader[1].ToString();
                        Items[k][5] = reader[2].ToString();
                        k++;
                        Items.Add(new string[6]);
                    }

                    reader.Close();
                    for (int l = 0; l < k; l++)
                    {
                        Product_query = String.Format("SELECT Item_name, Id_category,Producer_name FROM Items WHERE Id_item = {0}", Items[l][0]);
                        command = new OleDbCommand(Product_query, myConnection);
                        reader = command.ExecuteReader();
                        categ = "";
                        while (reader.Read())
                        {
                            Items[l][1] = reader[0].ToString();
                            Items[l][3] = reader[2].ToString();
                            categ = reader[1].ToString();
                        }
                        reader.Close();
                        Product_query = String.Format("SELECT Category_name FROM Item_category WHERE Id_category = {0}", categ);
                        command = new OleDbCommand(Product_query, myConnection);
                        reader = command.ExecuteReader();
                        while (reader.Read())
                            Items[l][2] = reader[0].ToString();
                        reader.Close();
                        richTextBox1.Text += Items[l][1] + " " + Items[l][2] + " " + Items[l][3] + " " + Items[l][4] + " " + Items[l][5] + "\n";
                    }
                    Items_c = new List<string[]>();
                    int h = 0;
                    for (int index = 0; index < Items.Count; index++)
                    {
                        if (Category.Equals(Items[index][2]))
                        {
                            Items_c.Add(new string[6]);
                            for (int j = 0; j < 6;j++)
                            Items_c[h][j] = Items[index][j];
                            h++;
                        }
                    }
                    textBox1.Text += Items_c.Count;
                    Word.Table _table = word_doc.Tables.Add(para[ml].Range, Items_c.Count+1, 5, ref missing, ref missing);
                    _table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
                    _table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
                    _currentRange = _table.Cell(1, 1).Range;
                    _table.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorOrange;
                    _currentRange.InsertAfter("Название товара");
                    _currentRange = _table.Cell(1, 2).Range;
                    _table.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorPaleBlue;
                    _currentRange.InsertAfter("Категория");
                    _currentRange = _table.Cell(1, 3).Range;
                    _table.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 3).Shading.BackgroundPatternColor = Word.WdColor.wdColorSeaGreen;
                    _currentRange.InsertAfter("Изготовитель");
                    _currentRange = _table.Cell(1, 4).Range;
                    _table.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 4).Shading.BackgroundPatternColor = Word.WdColor.wdColorViolet;
                    _currentRange.InsertAfter("Количество товара");
                    _currentRange = _table.Cell(1, 5).Range;
                    _table.Cell(1, 5).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    _table.Cell(1, 5).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow;
                    _currentRange.InsertAfter("Цена");
                    for (int index = 0; index < Items_c.Count; index++)
                    {
                        for (int j = 1; j <= 5; j++)
                        {
                            _currentRange = _table.Cell(index+2, j).Range;
                            _table.Cell(index+2, j).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                            _currentRange.InsertAfter(Items_c[index][j]);
                        }
                    }
                    para[ml].Range.InsertParagraphAfter();
                    para.Add(word_doc.Paragraphs.Add(ref missing));
                    ml++;
                    para[ml].Range.InsertParagraphAfter();
                    para[ml].Range.GoToEditableRange();
                    _currentRange = para[ml].Range;
                }

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.DefaultExt = "doc";
                sfd.Filter = "Word files (*.doc)|*.doc|All files (*.*)|*.*";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    word_doc.SaveAs2(sfd.FileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
ref missing, ref missing, ref missing, ref missing,
ref missing, ref missing, ref missing, ref missing,
 ref missing);
                }
                // Закрыть.
                object save_changes = false;
                word_doc.Close(ref save_changes, ref missing, ref missing);
                word_app.Quit(ref save_changes, ref missing, ref missing);
                richTextBox1.Text += "Готово\n";

            }
        
    }
        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Товары
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //Торговая сеть
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //Торговая точка
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Информация по торговым точкам
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //Логин
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            //Пароль
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            //Поле для лог-сообщений программы
        }
    }
}
