using System;
using System.Drawing;
using Spire.Doc.Fields;
using Spire.Doc;
using Spire.Doc.Documents;

namespace aula7_manipulacao_word
{
    class Program
    {
        static void Main(string[] args)
        {
            #region criacao do documento
                // criar documento com o nome exemploDoc
            Document exemploDoc                  =    new Document();
            #endregion 
               //adiciona uma secao com o nome secaoCapa ao documento
               //cada secao pode ser entendida como uma pagina do documento
            #region criacao de secao no documento
            Section secaoCapa = exemploDoc.AddSection();
            #endregion    
               //criar um paragrafo com nome titulo e adiciona a secao secaoCapa
               //os paragrafos sao necessarios para a insercao de textos, imagens, tabelas e etc
            #region criar um paragrafo
            Paragraph titulo                     =    secaoCapa.AddParagraph();
            #endregion
            #region adiciona texto ao paragrafo
                //adiciona o texto  exemplo de titulo ao paragrafo titulo
            titulo.AppendText("Exemplo de título \n\n");
            #endregion
            #region formatar paragrafo
                //atraves de propiedade horizontalalignment e possivel alinhar o paragrafo
            titulo.Format.HorizontalAlignment     =    HorizontalAlignment.Center;
             ParagraphStyle estilo01              =    new ParagraphStyle(exemploDoc);
                //adiciona um nome ao estilo
             estilo01.Name                        =    "cor do titulo";
                //defineir a cor do titulo
             estilo01.CharacterFormat.TextColor   =    Color.DarkBlue;  
                //define q o texto sera negrito
             estilo01.CharacterFormat.Bold        =    true;
                //adiciona o estilo01 ao documento exemploDoc
             exemploDoc.Styles.Add(estilo01);
                //aplica o estilo01 ao paragrafo titulo
             titulo.ApplyStyle(estilo01.Name);
             
            #endregion
            #region trabalhar com tabulaco
                //adiciona um paragrafo textoCapa a secao  secaoCapa
            Paragraph textoCapa                   =    secaoCapa.AddParagraph();
                //adiciona um texto ao paragrafo cm tabulacao
            textoCapa.AppendText("\teste e um exemplo de texto com tabulacao\n");
            Paragraph textoCapa2                  =    secaoCapa.AddParagraph();
               //adiciona umtexto ao paragrafo textocapa2 cm conctenacao/
            textoCapa2.AppendText("\tbasicamente,entao, uma secao representa uma pagina do documento a os paragrafos dentro de uma mesma secao,"+"obviamente, aparecem na mesmapagina.");
            #endregion
            #region 
               // adiciona um paragrafo a secao secaoCapa
            Paragraph imagemCapa                  =    secaoCapa.AddParagraph();
               //centralizar horizontalmente o paragrafo imagemCapa
            imagemCapa.AppendText("\n\n\tAgora vamos inserir uma imagem ao documento\n\n");

            imagemCapa.Format.HorizontalAlignment=     HorizontalAlignment.Center;
            DocPicture imagemExemplo              =    imagemCapa.AppendPicture(Image.FromFile(@"saida\imagem\logo_csharp.png"));
               //definir largura e altura
            imagemExemplo.Width                   =    300;
            imagemExemplo.Height                  =    300;
            #endregion
            #region
               //adiciona uma nova secao
            Section secaoCorpo                    =    exemploDoc.AddSection();
               //adiciona um paragrafo a secao secaoCorpo
            Paragraph paragraphCorpo1             =    secaoCapa.AddParagraph();
            paragraphCorpo1.AppendText("\teste e um exemplo de paragrafos criado em uma nova secao."+"\tcomo foi criada uam nova secao,perceba que esta texto aparecer em uma nova pagina.");

            #endregion
            #region adicionar uma tabela
            //adiciona uma tabela a secao secaocorpo
            Table tabela                          =    secaoCorpo.AddTable(true);
            //cria o cabecalho da tabela
            String[] cabecalho={"item","descrição","qtd.","preço Unit.","preço"};
            //criar dados da tabela
            String[][] dados                      ={
                new String[]{"cenoura", "vegetal muito nutritivo" ,"1", "R$ 4,00","R$ 4,00",},
                new String[]{"batata", "vegetal muito nutritivo" ,"", "R$ 5,00","R$ 10,00",},
                new String[]{"alface", "vegetal muito nutritivo" ,"1", "R$ 1,50","R$ 1,50",},
                new String[]{"tomate", "tomate e uma fruta " ,"2", "R$ 6,00","R$ 12,00",},

            };
            //adiciona as celulas na tabela
            tabela.ResetCells(dados.Length        +    1,cabecalho.Length);
            //adiciona uma linha na posicao [0] do vetor de linhas
            TableRow linha1                       =    tabela.Rows[0];
            linha1.IsHeader                       =    true;

            //define a altura d linha
            linha1.Height                         =23;
            //formatacao do cabecalho
            linha1 .RowFormat.BackColor           =    Color.AliceBlue;
            //percorre as colunas do cabecalho
            for (int i = 0; i < cabecalho.Length ; i++)
            {
            Paragraph p                           =    linha1.Cells[i].AddParagraph();
            linha1.Cells[i].CellFormat.VerticalAlignment= VerticalAlignment.Middle;
            p.Format.HorizontalAlignment          =    HorizontalAlignment.Center;  

            //formatacao dos dados do cabecalho
            TextRange TR                          =    p.AppendText(cabecalho[i]);
            TR.CharacterFormat.FontName           =    "calilbri";
            TR.CharacterFormat.FontSize           = 14;
            TR.CharacterFormat.TextColor          =    Color.Teal;
            TR.CharacterFormat.Bold               =    true;
            }
            //adiciona as linhas do corpo da tabela
            for (int r = 0; r < dados.Length; r++)
            {
                TableRow linhaDados               =    tabela.Rows[r + 1];
                //define a altura da linha
                linhaDados.Height                 = 20;
               // percorre as colunas
                for (int c = 0; c < dados[r].Length; c++)
                {
            //alinha celulas
            linhaDados.Cells[c].CellFormat.VerticalAlignment=VerticalAlignment.Middle;
            //preencher os dados da linha
            Paragraph p2                          =    linhaDados.Cells[c].AddParagraph();
            TextRange TR2                         =    p2.AppendText(dados[r][c]);
            //formata as celulas
            p2.Format.HorizontalAlignment         =    HorizontalAlignment.Center;
            TR2.CharacterFormat.FontName          =    "calibri";
            TR2.CharacterFormat.FontSize          =    12;
            TR2.CharacterFormat.TextColor         =    Color.Brown;
                }
            }
            #endregion



            #region salvar
            //salva o arquivo em .Docx
            //utiliza o metodo savetofile para salvar p arquivo no formato desejado
            //assim como no word , caso ja exista um arquivo com este nome , e substituido
               exemploDoc.SaveToFile(@"saida\exemplo_arquivo_word.docx", FileFormat.Docx);
            #endregion
            




            // []--->vetor
        }
    }
}
