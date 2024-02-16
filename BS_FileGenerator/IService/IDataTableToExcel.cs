using System.Data;

namespace BS_FileGenerator.IService
{
    public interface IDataTableToExcel
    {
        void ToExcel(DataTable table, string filepath, string sheetName = null);
    }
}