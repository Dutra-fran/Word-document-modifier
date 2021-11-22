using System;
using Xceed.Words.NET;
using ProjetoWordIntegrado_CSharp.Models;
using System.IO;

namespace ProjetoWordIntegrado_CSharp.Modifier_Word
{
    // Essa é a nossa classe que efetuará a ação de transcrever os valores nos campos do arquivo Word que o usuário deseja.
    public class Replace
    {
        public static void ReplaceElements(string _Atendente, int _Brinde, string _Nome1, string _Nome2, string _Cpf1, string _Cpf2, string _Contato1, string _Contato2)
        {
            // Pegando o diretório em que nosso programa situa-se.
            var DiretorioAtual = Directory.GetCurrentDirectory();

            // Inicializando nosso arquivo docx.
            using (var documento = DocX.Load(DiretorioAtual + @"\Documentos\" + "Template.docx"))
            {
                // Instanciando nossa classe model com a finalidade de pegar o horário atual da modificação. (Conforme usuário pediu).
                Cadastro hora = new Cadastro();

                // A seguir serão feitas algumas modificações, de fato, em nosso arquivo docx. É um processo de 'replace' devido a troca das palavras-chave pelo valor informado do usuário.
                documento.ReplaceText("#Atendente", _Atendente);

                switch (_Brinde)
                {
                    case 1:
                        documento.ReplaceText("#brinde", "Chocolate");
                        break;
                    case 2:
                        documento.ReplaceText("#brinde", "Bis");
                        break;
                    case 3:
                        documento.ReplaceText("#brinde", "Tortuguita");
                        break;
                }

                documento.ReplaceText("#DataAtendimento", hora.DataDeAtendimento.ToString());
                documento.ReplaceText("#nome1", _Nome1);
                documento.ReplaceText("#nome2", _Nome2);
                documento.ReplaceText("#cpf1", _Cpf1);
                documento.ReplaceText("#cpf2", _Cpf2);
                documento.ReplaceText("#contato1", _Contato1);
                documento.ReplaceText("#contato2", _Contato2);

                // A seguir foi feita uma implementação para ver se tal arquivo já existe, para evitar que eles sejam sobrescritos.
                string curFile = DiretorioAtual + @"\Documentos\Salvos\" + "novo-documento.docx";
                if (File.Exists(curFile))
                {
                    int i = 1;
                    while (true)
                    {
                        curFile = DiretorioAtual + @"\Documentos\Salvos\" + "novo-documento-" + i.ToString() + ".docx";
                        if (!File.Exists(curFile))
                        {
                            documento.SaveAs(DiretorioAtual + @"\Documentos\Salvos\" + "novo-documento-" + i.ToString() + ".docx");
                            break;
                        }
                        i++;
                    }
                }
                else
                    documento.SaveAs(DiretorioAtual + @"\Documentos\Salvos\" + "novo-documento.docx");

                Console.Clear();
            }
        }
    }
}