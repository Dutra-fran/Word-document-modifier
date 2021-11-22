using System;
using ProjetoWordIntegrado_CSharp.Models;
using System.IO;
using ProjetoWordIntegrado_CSharp.Modifier_Word;

namespace ProjetoWordIntegrado_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            bool verifica; // Variável para verificar se o usuário quer continuar no programa.
            char caractereConfirma; // Variável que guarda o valor passado do teclado do usuário se ele deseja continuar no programa ou não.
            var DiretorioAtual = Directory.GetCurrentDirectory(); // Pegando o caminho da pasta atual do programa.
            Cadastro _cadastro = new Cadastro(); // instanciando a nossa model, é nela que contém as informações que o usuário deseja passar ao documento Word.
            Console.WriteLine("Caminho atual da pasta: " + DiretorioAtual);

            do
            {
                Console.Write("Insira o nome do atendente: ");
                _cadastro.Atendente = Console.ReadLine();

                Console.Write("Brinde: Selecione uma opcao\n\t1. Chocolate\n\t2. Bis\n\t3. Tortuguita\n");
                _cadastro.Brinde = int.Parse(Console.ReadLine());

                Console.Write("Insira nome1: ");
                _cadastro.Nome1 = Console.ReadLine();

                Console.Write("Insira nome2: ");
                _cadastro.Nome2 = Console.ReadLine();

                do {
                    Console.Write("Insira o CPF-1: ");
                    _cadastro.Cpf1 = Console.ReadLine();
                } while (_cadastro.Cpf1 == null);

                do
                {
                    Console.Write("Insira o CPF-2: ");
                    _cadastro.Cpf2 = Console.ReadLine();
                } while (_cadastro.Cpf2 == null);

                Console.Write("Insira o contato1: ");
                _cadastro.Contato1 = Console.ReadLine();

                Console.Write("Insira o contato2: ");
                _cadastro.Contato2 = Console.ReadLine();

                Replace.ReplaceElements(_cadastro.Atendente, _cadastro.Brinde, _cadastro.Nome1, _cadastro.Nome2, _cadastro.Cpf1, _cadastro.Cpf2, _cadastro.Contato1, _cadastro.Contato2);
                Console.WriteLine("Modificações realizadas com sucesso!");

                // Laço while para fazer com que o programa persista ou não, depende do que o usuário passar.
                while (true)
                {
                    Console.Write("Você deseja continuar no programa? [N/s]: ");
                    caractereConfirma = Char.Parse(Console.ReadLine());

                    if (caractereConfirma == 'N' || caractereConfirma == 'n')
                    {
                        verifica = false;
                        break;
                    } else if(caractereConfirma == 'S' || caractereConfirma == 's')
                    {
                        verifica = true;
                        Console.Clear();
                        break;
                    } else
                    {
                        Console.WriteLine("Insira um valor válido!");
                    }
                }
            } while (verifica);
        }
    }
}