using envioMasivoCorreos.Dtos;

namespace envioMasivoCorreos.Utils
{
    public class Util
    {
        public IEnumerable<string> GetData()
        {
            var contents = File.ReadAllText("C:/excelPruebas.csv").Split('\n');
            return contents;
        }
    }
}