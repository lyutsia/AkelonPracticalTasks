using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;

namespace Task3
{
    public class DataFromExcel
    {
        private const string CodePropertyName = "Code";
        private WareData _wareData;
        private OrderData _orderData;
        private ClientData _clientData;
        private List<IData> _dataList;

        public string PathFile { get; set; }
        public DataFromExcel()
        {
            _wareData = new WareData();
            _orderData = new OrderData();
            _clientData = new ClientData();
            _dataList = new List<IData> { _wareData, _orderData, _clientData };
        }

        public DataFromExcel(string pathFile) : this()
        {
            PathFile = pathFile;
        }

        #region private methods
        /// <summary>
        /// заполнение заголовка
        /// </summary>
        /// <param name="row"></param>
        /// <param name="workbookPart"></param>
        /// <param name="data"></param>
        private void SetHeaders(Row row, WorkbookPart workbookPart, IData data)
        {
            foreach (var cell in row.Descendants<Cell>())
            {
                data.Headers.Add(GetCellValue(cell, workbookPart), GetColumnIndexByCellReference(cell.CellReference));
            }
        }

        /// <summary>
        /// создание объекта по типу
        /// </summary>
        /// <param name="type"></param>
        /// <param name="row"></param>
        /// <param name="workbookPart"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        private object CreateObjectByType(Type type, Row row, WorkbookPart workbookPart, IData data)
        {
            var properties = type.GetProperties();
            var propertyCode = properties.FirstOrDefault(p => p.Name == CodePropertyName);
            var cellValueByPropertyCode = GetCellValueByProperty(propertyCode, row, data, workbookPart, out Type typeProp);
            var newObject = GetObjectByCode(type, cellValueByPropertyCode, data);
            if (newObject == null)
                newObject = Activator.CreateInstance(type);

            foreach (var property in type.GetProperties())
            {
                if (property.Name == CodePropertyName)
                    continue;

                var cellValueByProperty = GetCellValueByProperty(property, row, data, workbookPart, out Type typeProperty);
                if (cellValueByProperty == null)
                    continue;

                SetPropertyValue(property, typeProperty, newObject, cellValueByProperty);
            }

            return newObject;
        }

        /// <summary>
        /// возвращает значение ячейки по свойству
        /// </summary>
        /// <param name="property"></param>
        /// <param name="row"></param>
        /// <param name="data"></param>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private string GetCellValueByProperty(PropertyInfo property, Row row, IData data, WorkbookPart workbookPart, out Type typeProperty)
        {
            typeProperty = null;
            var columnCaptionAttributeFromProperty = property.GetCustomAttribute(typeof(ColumnCaptionAttribute), false) as ColumnCaptionAttribute;
            if (columnCaptionAttributeFromProperty == null)
                return null;

            var columnCaptionFromProperty = columnCaptionAttributeFromProperty.Caption;
            if (!data.Headers.ContainsKey(columnCaptionFromProperty))
                return null;

            typeProperty = columnCaptionAttributeFromProperty.TypeProperty;
            var cell = row.Descendants<Cell>().FirstOrDefault(cell => GetColumnIndexByCellReference(cell.CellReference) == data.Headers[columnCaptionFromProperty]);
            return GetCellValue(cell, workbookPart);
        }

        /// <summary>
        /// установка значения свойству
        /// </summary>
        /// <param name="property"></param>
        /// <param name="type"></param>
        /// <param name="obj"></param>
        /// <param name="value"></param>
        private void SetPropertyValue(PropertyInfo property, Type type, object obj, object value)
        {
            if (type == typeof(int))
            {
                int.TryParse(value.ToString(), out int intValue);
                property.SetValue(obj, intValue);
            }
            else if (type == typeof(decimal))
            {
                decimal.TryParse(value.ToString(), out decimal decimalValue);
                property.SetValue(obj, decimalValue);
            }
            else if (type == typeof(DateTime))
            {
                double.TryParse(value.ToString(), out double dateDooble);
                var dateTimeValue = DateTime.FromOADate(dateDooble);
                property.SetValue(obj, dateTimeValue);
            }
            else
            {
                var data = _dataList.FirstOrDefault(d => d.ObjectType == type);
                if (data != null)
                    value = GetObjectByCode(type, value, data);
                property.SetValue(obj, value);
            }
        }
        /// <summary>
        /// установка ссылочного значения свойству
        /// </summary>
        /// <param name="property"></param>
        /// <param name="type"></param>
        /// <param name="value"></param>
        /// <param name="data"></param>
        private object GetObjectByCode(Type type, object value, IData data)
        {
            int.TryParse(value.ToString(), out int code);
            var obj = data.GetObjectByCode(code);
            if (obj == null && code != 0)
            {
                obj = Activator.CreateInstance(type, code);
                data.AddObject(obj);
            }
            return obj;
        }

        /// <summary>
        /// получение значения из ячейки
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            var value = cell.InnerText;
            if (cell.DataType == null)
            {
                return value;
            }
            if (cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                if (stringTable != null)
                {
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }
        /// <summary>
        /// получить индекс колонки в эксель по ячейке
        /// </summary>
        /// <returns></returns>
        private string GetColumnIndexByCellReference(StringValue stringValue)
        {
            var columnIndex = string.Empty;
            foreach (var ch in stringValue.Value)
            {
                if (char.IsLetter(ch))
                    columnIndex += ch;
            }
            return columnIndex;
        }
        /// <summary>
        /// сохраняем изменение свойства объекта в файле
        /// </summary>
        /// <param name="codeObj"></param>
        /// <param name="codePropertyIndexColumn"></param>
        /// <param name="changePropertyValue"></param>
        /// <param name="changePropertyIndexColumn"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        private bool SaveChangePropertyInExcel(int codeObj, string codePropertyIndexColumn, string changePropertyValue, string changePropertyIndexColumn, IData data)
        {
            try
            {
                using (var document = SpreadsheetDocument.Open(PathFile, true))
                {
                    var sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().SingleOrDefault(s => s.Name == data.SheetName);
                    var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);

                    foreach (var row in worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>())
                    {
                        var cell = row.Descendants<Cell>().FirstOrDefault(cell => GetColumnIndexByCellReference(cell.CellReference) == codePropertyIndexColumn);
                        var cellValue = GetCellValue(cell, document.WorkbookPart);
                        int.TryParse(cellValue, out int code);
                        if (code == codeObj)
                        {
                            cell = row.Descendants<Cell>().FirstOrDefault(cell => GetColumnIndexByCellReference(cell.CellReference) == changePropertyIndexColumn);
                            cell.CellValue = new CellValue(changePropertyValue);
                            break;
                        }
                    }
                    document.WorkbookPart.Workbook.Save();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region public methods для работы с данными из файла

        /// <summary>
        /// загрузка данных из xlsx-файла
        /// </summary>
        /// <param name="pathFile"></param>
        public bool LoadDataFromExcel(string pathFile)
        {
            try
            {
                using (var document = SpreadsheetDocument.Open(pathFile, false))
                {
                    foreach (var data in new List<IData> { _wareData, _orderData, _clientData })
                    {
                        var sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().SingleOrDefault(s => s.Name == data.SheetName);
                        var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);

                        var firstRow = false;
                        foreach (var row in worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>())
                        {
                            if (!firstRow)
                            {
                                SetHeaders(row, document.WorkbookPart, data);
                                firstRow = true;
                                continue;
                            }
                            CreateObjectByType(data.ObjectType, row, document.WorkbookPart, data);
                        }
                    }
                }
                return true;
            }
            catch
            {
                Console.WriteLine("Не удалось загрузить данные!");
                return false;
            }
        }
        /// <summary>
        /// получение информации о клиенте, который заказывал товар с заданным названием
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        public string GetInfoAboutClientsOrderedWareByCaption(string caption)
        {
            var findWares = _wareData.GetWareByCaption(caption);
            if (findWares.Count() == 0)
                return "Нет товаров с заданным названием!";

            var orders = _orderData.GetOrdersWithFindWares(findWares);
            var outputText = "Клиенты, заказавшие этот товар:\n";
            foreach (var order in orders)
            {
                outputText += $"{order.Client.ToString()} \tКол-во: {order.RequiredQuantity} \tЦена: {order.Ware.Price} \tДата: {order.DatePosting.ToShortDateString()}\n";
            }
            return outputText;

        }
        /// <summary>
        /// получение клиента с наибольшим количеством заказов за месяц
        /// </summary>
        /// <returns></returns>
        public IClient GetClientWithMaxOrdersInMonth(DateTime dateTime) => _orderData.GetClientWithMaxOrdersInMonth(dateTime);

        /// <summary>
        /// получение клиента с наибольшим количеством заказов за год
        /// </summary>
        /// <returns></returns>
        public IClient GetClientWithMaxOrdersInYear(DateTime dateTime) => _orderData.GetClientWithMaxOrdersInYear(dateTime);

        /// <summary>
        /// изменение контактного лица клиента у заданной компании
        /// </summary>
        /// <returns></returns>
        public bool SetContactPersonNameByCompany(string company, string newContactPersonName)
        {
            var client = _clientData.GetClientByCompany(company);
            if (client == null)
            {
                Console.WriteLine("Клиент не найден!");
                return false;
            }
            client.ContactPersonName = newContactPersonName;

            //получаем индексы колонок по нужных свойств
            var properties = typeof(Client).GetProperties();
            var codePropertyColumnCaptionAttribute = properties.FirstOrDefault(prop => prop.Name == nameof(client.Code))
                                                               .GetCustomAttribute(typeof(ColumnCaptionAttribute), false) as ColumnCaptionAttribute;
            var contactPersonNamePropertyColumnAttribute = properties.FirstOrDefault(prop => prop.Name == nameof(client.ContactPersonName))
                                                                    .GetCustomAttribute(typeof(ColumnCaptionAttribute), false) as ColumnCaptionAttribute;

            var codePropertyIndexColumn = _clientData.Headers[codePropertyColumnCaptionAttribute.Caption];
            var contactPersonNamePropertyIndexColumn = _clientData.Headers[contactPersonNamePropertyColumnAttribute.Caption];

            return SaveChangePropertyInExcel(client.Code, codePropertyIndexColumn, client.ContactPersonName, contactPersonNamePropertyIndexColumn, _clientData);
        }
        #endregion
    }
}
