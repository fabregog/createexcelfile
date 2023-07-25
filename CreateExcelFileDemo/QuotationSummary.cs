using System.Text.Json.Serialization;

namespace CreateExcelFileDemo;
public class QuotationSummary
{
    [JsonPropertyName("id")]
    public int Id { get; set; }
    [JsonPropertyName("customerName")]
    public string CustomerName { get; set; }
    [JsonPropertyName("carBrand")]
    public string CarBrand { get; set; }
    [JsonPropertyName("carModel")]
    public string CarModel { get; set; }
    [JsonPropertyName("anualPayment")]
    public decimal AnualPayment { get; set; }
}
