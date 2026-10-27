
# WE2002 Home Team Editor ML v0.1

**Editor de Estádios da Master Liga do Winning Eleven 2002 (PS1)**

**Desenvolvido por Fabio FJA | Visual Basic 6.0 | Windows | PlayStation 1**

## Sobre o projeto

O WE2002 Home Team Editor ML v0.1 é uma ferramenta de edição e personalização do Winning Eleven 2002 para PlayStation 1 (PSX), desenvolvida em Visual Basic 6.0.

O programa permite modificar o estádio utilizado como casa por cada um dos 32 clubes da Master Liga, oferecendo uma interface gráfica para selecionar equipes, visualizar estádios e salvar as alterações diretamente no arquivo BIN do jogo.

A ferramenta utiliza técnicas de engenharia reversa e manipulação binária, alterando offsets específicos para personalizar a configuração dos estádios sem a necessidade de utilizar um editor hexadecimal manualmente.

## Funcionalidades

### 1. Editor de estádios da Master Liga
- Modificação do estádio mandante dos clubes da Master Liga.
- Configuração individual de estádio para cada uma das 32 equipes.
- Seleção entre 16 identificadores internos de estádios.
- Leitura das configurações existentes no arquivo BIN.
- Gravação das alterações diretamente no jogo.
- Possibilidade de alterar novamente os estádios já configurados.

### 2. Suporte a 32 clubes

O editor permite modificar os estádios dos seguintes clubes:

**Inglaterra**
- Manchester United
- Arsenal
- Chelsea
- Liverpool
- Leeds
- Newcastle
- Aston Villa

**Espanha**
- Barcelona
- Real Madrid
- Valencia
- Deportivo La Coruña

**França**
- Monaco
- Marseille
- Paris Saint-Germain
- Bordeaux

**Holanda**
- Ajax
- Feyenoord
- PSV

**Itália**
- Internazionale
- Juventus
- Milan
- Lazio
- Parma
- Fiorentina
- Roma

**Alemanha**
- Borussia Dortmund
- Bayern Munich
- Bayer Leverkusen

**Outros clubes**
- Olympiacos
- Galatasaray
- Dynamo Kyiv
- Boca Juniors

### 3. Seleção de 16 estádios

O programa oferece 16 opções de estádios identificadas pelos códigos internos:

`GDC_DD`, `GDC_FD`, `GDC_JD`, `GDC_SD`, `GDC_MD`, `GDC_AD`, `GDC_PD`, `GDC_CD`, `GDC_HD`, `GDC_TD`, `GDC_GDJ`, `GDC_MDJ`, `GDC_RDJ`, `GDC_ID`, `GDC_OD` e `GDC_VD`.

Cada opção corresponde a um valor binário utilizado pelo jogo para determinar o estádio associado à equipe.

### 4. Pré-visualização gráfica
- Exibição de imagens representativas dos estádios.
- Atualização da imagem ao selecionar um estádio.
- Visualização do estádio associado à configuração lida do jogo.
- Interface com recursos gráficos incorporados aos formulários.

### 5. Leitura e edição de arquivos BIN
- Seleção do arquivo BIN por uma janela do Windows.
- Leitura automática do valor correspondente ao clube selecionado.
- Identificação do estádio atualmente configurado.
- Gravação direta de novos valores binários.
- Mensagem de confirmação após salvar.

### 6. Documentação de offsets

O projeto inclui a planilha:

`ORDEM DOS ESTADIOS NA MASTER by Fabio FJA.xlsx`

Ela documenta a ordem das equipes, posições binárias, códigos de estádios e seus respectivos valores hexadecimais e decimais.

Esse material auxilia na compreensão das estruturas internas utilizadas pelo Winning Eleven 2002.

## Tecnologias utilizadas

- **Linguagem:** Visual Basic 6.0
- **Interface gráfica:** Visual Basic Forms
- **Ambiente:** Microsoft Visual Basic 6
- **Manipulação binária:** comandos Open, Get e Put
- **Componentes:** Microsoft Common Dialog Control (COMDLG32.OCX)
- **Recursos gráficos:** BMP, JPG e FRX
- **Plataforma:** Microsoft Windows
- **Jogo:** Winning Eleven 2002 (PlayStation 1)

## Como funciona

O editor utiliza uma tabela de 32 offsets consecutivos, localizados entre os endereços decimais 2833127 e 2833158 do arquivo BIN.

Cada posição corresponde a uma equipe da Master Liga e armazena um valor que identifica seu estádio mandante.

Ao selecionar um clube, o programa lê o byte correspondente, identifica o estádio e apresenta sua representação gráfica.

Quando o usuário escolhe outro estádio e utiliza a opção de salvar, o programa substitui o valor diretamente na imagem BIN.

## Como utilizar

1. Execute `Home Team Editor ML v_0.1 by fja.exe`.
2. Selecione o arquivo BIN do Winning Eleven 2002 na janela de abertura.
3. Escolha uma equipe no campo **Equipe**.
4. Observe o estádio atualmente configurado no campo **Casa**.
5. Selecione um novo estádio.
6. Confira sua pré-visualização gráfica.
7. Clique em **Salvar** para aplicar a alteração.
8. Repita o processo para os demais clubes desejados.

## Estrutura do projeto

- `Form1.frm`: formulário principal e lógica do editor.
- `Form1.frx`: recursos gráficos da interface principal.
- `Form2.frm`: janela de informações e créditos.
- `Form2.frx`: recursos da janela Sobre.
- `Project1.vbp`: projeto original do Visual Basic 6.
- `Bandeiras/`: imagens de bandeiras das equipes.
- `Stadiums/`: imagens ilustrativas dos estádios.
- `Recursos/`: imagens e ícones utilizados no projeto.
- `ORDEM DOS ESTADIOS NA MASTER by Fabio FJA.xlsx`: documentação dos offsets e identificadores de estádios.
- `Home Team Editor ML v_0.1 by fja.exe`: programa executável para Windows.

## Observações importantes

- Faça um backup do BIN original antes de utilizar o editor.
- As modificações são gravadas diretamente no arquivo selecionado.
- O programa utiliza offsets específicos do Winning Eleven 2002.
- Não existe validação automática da versão da imagem BIN.
- Algumas funcionalidades podem apresentar limitações por se tratar de uma versão experimental.

## Créditos

**Desenvolvimento e programação:** Fabio FJA

**Versão:** 0.1

**Linguagem:** Visual Basic 6.0

**Jogo:** Winning Eleven 2002

**Plataforma:** PlayStation 1

Projeto independente desenvolvido para estudo, pesquisa, engenharia reversa e personalização de jogos clássicos de futebol.

Este projeto não possui vínculo oficial com a Konami.
