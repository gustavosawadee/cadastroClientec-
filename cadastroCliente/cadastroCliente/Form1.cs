using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Data.OleDb; 

namespace cadastroCliente
{
    public partial class Formulario : Form
    {
        //string de conexão ATENÇÂO !!! substituir \ por \\ 
        static string strCn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\Aluno_2\\Downloads\\cliente.accdb";

        OleDbConnection conexao = new OleDbConnection(strCn);

        public Formulario()

        {
            InitializeComponent();
        }

        private void btn_consulta_Click(object sender, EventArgs e)
        {
            //instrução sql responsável por pesquisar o banco de dados (CRUD - Read)
            string pesquisa = "select * from cadastroCliente where Id = " + txt_id.Text;
            //criando um objeto de nome cmd tendo como modelo a classe OleDbCommand para executar a instrução sql
            OleDbCommand cmd = new OleDbCommand(pesquisa, conexao);
            // Atravé da classe OleDbDataReader que faz parte do SqlCliente, criamos uma //variável chamada DR que será usada na leitura dos dados (instrução select)
            OleDbDataReader DR;
            //tratamento de exceções: try - catch - finally (em caso de erro capturamos o //tipo do erro)

try 
{
// Abrindo a conexão com o banco
conexao.Open();
// Executando a instrução e armazenando o resultado no reader DR
DR = cmd.ExecuteReader();
// Se houver um registro correspondente ao Id
if (DR.Read()) 

{ 
// Exibe as informações nas caixas de texto (textBox) correspondentes (0) //corresponde ao Id, (1) ao Nome e assim sucessivamente 
txt_id.Text = DR.GetValue(0).ToString(); 
txt_nome.Text = DR.GetValue(1).ToString(); 
txt_endereco.Text = DR.GetValue(2).ToString(); 
txt_cidade.Text = DR.GetValue(3).ToString();
txt_estado.Text = DR.GetValue(4).ToString();
txt_rg.Text = DR.GetValue(5).ToString();
txt_cpf.Text = DR.GetValue(6).ToString(); 
} 

// Senão, exibimos uma mensagem avisando e também limpamos os campos para uma //nova pesquisa 
else 
{ 
MessageBox.Show("Registro não encontrado"); 
txt_nome.Clear(); 
txt_endereco.Clear(); 
txt_cidade.Clear(); 
txt_estado.Clear();
txt_rg.Clear();
txt_cpf.Clear();
txt_id.Focus(); 

} // Encerrando o uso do reader DR 
DR.Close(); 

// Encerrando o uso do cmd 
cmd.Dispose(); 
} 

//caso ocorra algum erro 
catch (Exception ex) 
{
 
//exiba qual é o erro 
MessageBox.Show(ex.Message); 
} 

// de qualquer forma sempre fechar a conexão com o banco ("lembrar da porta da //geladeira rsrsrs") 
finally 
{ 
conexao.Close(); 
} 
}

        private void btn_salvar_Click(object sender, EventArgs e)
        {
            //instrução sql responsável por adicionar dados ao banco (CRUD - Create) 
string adiciona = "insert into cadastroCliente values (" + 
txt_id.Text + ",'" + 
txt_nome.Text + "','" + 
txt_endereco.Text + "','" +
txt_cidade.Text + "','" +
txt_estado.Text + "','" +
txt_rg.Text + "','" +
txt_cpf.Text + "')"; 

//criando um objeto de nome cmd tendo como modelo a classe OleDbCommand para //executar a instrução sql 
OleDbCommand cmd = new OleDbCommand(adiciona, conexao);
 
//tratamento de exceções: try - catch - finally (em caso de erro capturamos o //tipo do erro) 
try 
{ 
// Abrindo a conexão com o banco 
conexao.Open(); 

// Criando uma variável para adicionar e armazenar o resultado 
int resultado;
resultado = cmd.ExecuteNonQuery(); 

// Verificando se o registro foi adicionado 
// Caso o valor da variável resultado seja 1 
// significa que o comando funcionou, neste caso limpar os campos e exibir uma //mensagem 
if (resultado == 1) 
{ 
MessageBox.Show("Registro adicionado com sucesso"); 
txt_nome.Clear(); 
txt_endereco.Clear(); 
txt_cidade.Clear(); 
txt_estado.Clear();
txt_rg.Clear();
txt_cpf.Clear();
txt_id.Focus(); 
}
 
// Encerrando o uso do cmd 
cmd.Dispose(); 
}
 
//caso ocorra algum erro 
catch (Exception ex) 
{
 
//exiba qual é o erro 
MessageBox.Show(ex.Message); 
} 

// de qualquer forma sempre fechar a conexão com o banco ("lembrar da porta da //geladeira rsrsrs") 
finally 
{ 
conexao.Close(); 
} 
}

        private void btn_excluir_Click(object sender, EventArgs e)
        {
            //instrução sql responsável por remover um registro do banco (CRUD - Delete) 
string remove = "delete from cadastroCliente where Id= " + txt_id.Text; 

//criando um objeto de nome cmd tendo como modelo a classe OleDbCommand para //executar a instrução sql 
OleDbCommand cmd = new OleDbCommand(remove, conexao); 

//tratamento de exceções: try - catch - finally (em caso de erro capturamos o //tipo do erro) 
try 
{ 

// Abrindo a conexão com o banco 
conexao.Open(); 

// Criando uma variável para adicionar e armazenar o resultado 
int resultado; 
if (MessageBox.Show("Tem certeza que deseja remover este registro ?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) 
{ 
resultado = cmd.ExecuteNonQuery();
// Verificando se o registro foi apagado 
// Caso o valor da variável resultado seja 1 
// significa que o comando funcionou, neste caso limpar os campos e exibir uma //mensagem 
if (resultado == 1) 
{ 
txt_nome.Clear(); 
txt_endereco.Clear(); 
txt_cidade.Clear(); 
txt_estado.Clear();
txt_rg.Clear();
txt_cpf.Clear();
txt_id.Focus(); 
MessageBox.Show("Registro removido com sucesso"); 
} 

// Encerrando o uso do cmd 
cmd.Dispose(); 
} 
} 

//caso ocorra algum erro 
catch (Exception ex) 
{ 

//exiba qual é o erro 
MessageBox.Show(ex.Message); 
} 
// de qualquer forma sempre fechar a conexão com o banco 
finally 
{ 
conexao.Close(); 
} 
}

        private void btn_sair_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void btn_limpar_Click(object sender, EventArgs e)
        {
            txt_id.Clear();
            txt_nome.Clear();
            txt_endereco.Clear();
            txt_cidade.Clear();
            txt_estado.Clear();
            txt_rg.Clear();
            txt_cpf.Clear();}
    }
        }
