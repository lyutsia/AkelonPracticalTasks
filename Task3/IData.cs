namespace Task3
{
    public interface IData
    {
        Dictionary<string, string> Headers { get; set; }
        string SheetName { get; }

        Type ObjectType { get; }
        /// <summary>
        /// добавление объекта
        /// </summary>
        /// <param name="obj"></param>
        void AddObject(object obj);

        /// <summary>
        /// получение объекта по коду
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        object GetObjectByCode(int code);
    }
}
