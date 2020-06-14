using System.IO;

namespace ExchangeServerSubscriber
{
    public interface IWatermarkStorage
    {
        string Load();
        void Save(string watermark);
    }

    public class WatermarkStorage : IWatermarkStorage
    {
        private const string PathToFile = "watermark.txt";

        public string Load() => File.Exists(PathToFile) ? File.ReadAllText(PathToFile) : string.Empty;

        public void Save(string watermark) => File.WriteAllText(PathToFile, watermark);
    }
}