using OfficeOpenXml;
using System.Data.SqlClient;

namespace mtpbi.Repository
{
    public class LeituraRepository
    {
        // Caminho do diretório onde os arquivos serão lidos
        string diretorioEntrada = @"C:\Caminho\Para\Arquivos\Entrada";

        // Caminho do diretório onde os arquivos serão movidos após o processamento
        string diretorioSaida = @".\ arquivos processados\";

        // Conexão com o banco de dados
        string connectionString = "Data Source=SEUSERVIDOR;Initial Catalog=MultiStoreDW;Integrated Security=True;";
        SqlConnection conexao = new SqlConnection(connectionString);

        // Lista os arquivos no diretório de entrada
        string[] arquivos = Directory.GetFiles(diretorioEntrada, "*.xlsx");

            foreach (string arquivo in arquivos)
            {
                // Lê o arquivo xlsx
                using (ExcelPackage pacote = new ExcelPackage(new FileInfo(arquivo)))
                {
                    ExcelWorksheet planilha = pacote.Workbook.Worksheets[0];
        int totalLinhas = planilha.Dimension.End.Row;

        // Abre conexão com o banco de dados
        conexao.Open();

                    // Limpa os dados da tabela stage.MultiStore
                    SqlCommand limparTabela = new SqlCommand("TRUNCATE TABLE stage.MultiStore", conexao);
        limparTabela.ExecuteNonQuery();

                    // Insere os dados do arquivo na tabela stage.MultiStore
                    for (int i = 2; i <= totalLinhas; i++)
                    {
                        SqlCommand inserirDados = new SqlCommand(
                            "INSERT INTO stage.MultiStore (VendaID, DataVenda, ProdutoID, Quantidade, ValorTotal, NomeArquivo) " +
                            "VALUES (@VendaID, @DataVenda, @ProdutoID, @Quantidade, @ValorTotal, @NomeArquivo)", conexao);

        inserirDados.Parameters.AddWithValue("@VendaID", planilha.Cells[i, 1].Value);
                        inserirDados.Parameters.AddWithValue("@DataVenda", planilha.Cells[i, 2].Value);
                        inserirDados.Parameters.AddWithValue("@ProdutoID", planilha.Cells[i, 3].Value);
                        inserirDados.Parameters.AddWithValue("@Quantidade", planilha.Cells[i, 4].Value);
                        inserirDados.Parameters.AddWithValue("@ValorTotal", planilha.Cells[i, 5].Value);
                        inserirDados.Parameters.AddWithValue("@NomeArquivo", Path.GetFileName(arquivo));

                        inserirDados.ExecuteNonQuery();
                    }

                    // Fecha conexão com o banco de dados
                    conexao.Close();
                                }

                // Move o arquivo processado para o diretório de saída
                File.Move(arquivo, Path.Combine(diretorioSaida, Path.GetFileName(arquivo)));
     }
    }
}
