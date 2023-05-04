
<img src="https://user-images.githubusercontent.com/12186574/236285667-70239785-05e2-4889-965f-f1b579a0dc2d.png" width="200" >




Exercicio prático do parse na linguagem C#

Depois de ler o arquivo de planilha, o código obtém a primeira planilha usando a propriedade Workbook.Worksheets[0]. 

Em seguida, ele obtém o número total de linhas na planilha usando a propriedade Dimension.End.Row.

O código então cria uma lista vazia de objetos Person que serão preenchidos com os dados da planilha. 

Ele itera por cada linha na planilha, começando da segunda linha (a primeira linha contém os cabeçalhos da coluna) até a última linha. 

Para cada linha, o código lê os valores de cada célula usando a propriedade Cells[row, col].Value.ToString().

Converte os valores para o tipo C# desejado usando os métodos int.Parse() e double.Parse().

O código então cria um novo objeto Person com os valores lidos da planilha e adiciona esse objeto à lista de pessoas. 

Finalmente, o código itera pela lista de pessoas e exibe os dados de cada pessoa na saída padrão usando o método Console.WriteLine().
