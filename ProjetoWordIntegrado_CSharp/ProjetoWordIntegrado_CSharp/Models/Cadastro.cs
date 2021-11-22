using System;

namespace ProjetoWordIntegrado_CSharp.Models
{
    // Esta é a nossa classe model de Cadastro. Essas informações serão as informações que serão transcritas ao documento Word.
    public class Cadastro
    {
        private string _Cpf1;
        private string _Cpf2;

        public string Atendente { get; set; }
        public int Brinde { get; set; }
        public DateTime DataDeAtendimento { get => DateTime.Now; }
        public string Nome1 { get; set; }
        public string Nome2 { get; set; }
        public string Cpf1 {
            get => _Cpf1;
            set {
                if (value.Length == 14)
                    _Cpf1 = value;
            }
        }
        public string Cpf2 {
            get => _Cpf2;
            set
            {
                if (value.Length == 14)
                    _Cpf2 = value;
            }
        }
        public string Contato1 { get; set; }
        public string Contato2 { get; set; }
    }
}
