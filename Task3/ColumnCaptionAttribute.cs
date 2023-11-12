namespace Task3
{
    public class ColumnCaptionAttribute : Attribute
    {
        public ColumnCaptionAttribute(string caption, Type type)
        {
            Caption = caption;
            TypeProperty = type;
        }

        public string Caption { get; set; }
        public Type TypeProperty { get; set; }
    }
}
