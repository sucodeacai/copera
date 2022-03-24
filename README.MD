<h4 align="center"> 
	🚧 Em construção...  🚧
</h4>
<h1 align="center">COPERA - Construtor de Planilha de Gestão de Horários Acadêmicos</h1>


### 🚀1.Descrição do Projeto
<p align="left">A partir das necessidades da organização de horários acadêmicos foi desenvolvido uma planilha que permite realizar de forma fácil a gestão dessa atividade. </p>
<p align="left">
A planilha permite o controle de disciplinas, professores, carga horária por turmas, choques de horários, relatórios por professores e turmas dentre outras funcinalidades</p>
<p align="left"> 
Atualmente a planilha é composta por 5 abas públicas, sendo elas e seus repectivos campos: 
</p>


* Professores (Nome, Código), onde deve ser cadastrado todos os professores, em que o campo código deve ser um identicador único que será utilizado posteriormente por relatórios.

* Turmas (CódigoDasTurmas, Seleção de Tipo), todas as turmas geridas devem ser lançadas nessa planilha, preferencialmente o campo CódigoDasTurmas deve ser o nome da turma sem espaços, pois ele é utilizado para gerar as guias para cada turma e nos relatórios. O campo tipo está relacionado a modalidade do curso se é anual ou semestral, que implica diretamente no cálculo de encontros semanais.

* Disciplinas (Nome, Turma, Diciplina, C.H.), guia responsável por fazer relação entre as disciplinas, turma, carga horária e professores. Ela também contém os campos (Tipo, Encontros, Encontros Semanais e Aprsentação), que são gerados/calculados através das fórmulas.

* Relatório Professor (Seleção de Professor), por meio dessa guia é possível selecionar um professor e será gerado um relatório que permite checar a quantidade de encontros diários, semanais e se está acontecendo algum choque de horário.

* Relatório Turma (Seleção de Turma), por meio dessa guia é possível selecionar uma turma e verificar a carga horária total, quantidade de encontros semanais e total.

* O lançamento dos horários é feito na guia de cada turma, onde através de filtros são carregados suas respectivas disciplinas. E os horários são lançados através de lista de seleções.


### ✅ 2.Utilizando a solução  
* A solução foi utilizada e validada utilizando google sheets, logo todos os passos a seguir devem ser realizados utilizando ele.
* Abrir o arquivo copera.xlsx e atualizar as abas: Professores, Turmas, Disciplinas.
* Seguir intruções do item 3.
* Ao executar as ações, será gerado o arquivo copera_processed.xlsx, que pega as disciplinas e turmas cadastradas e gera suas repesctivas abas preenchidas. 

### 🎲 3.Utilizando o código
```bash
# Ambiente Linux
# Clone este repositório
git clone https://github.com/wramonbessa/copera.git
# Criar ambiente 
python3 -m venv venv
# Activate Python venv
source venv/bin/activate
#Atualiza pip
pip install --upgrade pip
# Intalar dependências
pip install -r requirements.txt
# Iniciar app
python3 app.py
```

### ✅ 4.Sobre
<p align="left"> 
A planilha surgiu inicialmente a partir de demanas da comissão de Horários do IFAC Campus Sena Madureira, e vem sendo constantemente atualizada pelos seus mantenedores. Dentre as novas funcionalidades pode-se citar a geração das fórmulas de forma automática por meio da linguagem python. 
</p>
<p align="left"> 
Devido a solicitações de compartilhamento foi desenvolvido esse repositório.
</p>

### ✅ 4.Mantenedores
* Ramon Bessa
* Jonas Pontes


