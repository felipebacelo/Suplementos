<p align="center">
<a href= "https://img.shields.io/github/repo-size/felipebacelo/Suplementos?style=for-the-badge"><img src="https://img.shields.io/github/repo-size/felipebacelo/Suplementos?style=for-the-badge"/></a>
<a href= "https://img.shields.io/github/languages/count/felipebacelo/Suplementos?style=for-the-badge"><img src="https://img.shields.io/github/languages/count/felipebacelo/Suplementos?style=for-the-badge"/></a>
<a href= "https://img.shields.io/github/forks/felipebacelo/Suplementos?style=for-the-badge"><img src="https://img.shields.io/github/forks/felipebacelo/Suplementos?style=for-the-badge"/></a>
<a href= "https://img.shields.io/bitbucket/pr-raw/felipebacelo/Suplementos?style=for-the-badge"><img src="https://img.shields.io/bitbucket/pr-raw/felipebacelo/Suplementos?style=for-the-badge"/></a>
<a href= "https://img.shields.io/bitbucket/issues/felipebacelo/Suplementos?style=for-the-badge"><img src="https://img.shields.io/bitbucket/issues/felipebacelo/Suplementos?style=for-the-badge"/></a>
</p>

# Suplementos
Suplementos em VBA Excel

Este repositório de Suplementos em VBA Excel foi desenvolvido com a finalidade de otimizar algumas tarefas rotineiras ao manusear planilhas. Novas funcionalidades serão integradas a este arquivo ao longo do tempo.

No arquivo __Suplementos__ estão até o momento as seguintes macros:

* Erros (Macro utilizada para colorir as células seleciondas com erros em vermelho)
* Maiúsculas (Macro utilizada para converter o valor das células selecionadas em maiúsculas)
* Minúsculas (Macro utilizada para converter o valor das células selecionadas em minúsculas)
***
### Requisitos

* Habilitar Macros
* Habilitar Guia de Desenvolvedor
***
### Referências às Bibliotecas

* Visual Basic For Applications
* Microsoft Excel 16.0 Object Library
* OLE Automation
* Microsoft Office 16.0 Object Library
***
### Compatibilidade

Este arquivo de Suplementos foi desenvolvido no Excel 2019 (64 bits).
Sua compatibilidade é garantida para a versão 2007 e superiores. Sua utilização em versões anteriores pode ocasionar em não funcionamento do mesmo.
***
### Usabilidade

Para utilizar os Suplementos o usuário deverá:

* Realizar o download do arquivo ZIP: __Suplementos__.
* Salvar o arquivo __Suplementos.xlam__ em qualquer pasta de trabalho de sua preferência.
* Abrir o Microsoft Excel.
* Habilitar as Macros e a Guia de Desenvolvedor.

O usuário poderá fixar o ícone de cada macro na _Barra de Ferramentas de Acesso Rápido_, para utilizar o suplemento.

* Clicar com o botão direito do mouse sobre a _Barra de Ferramentas de Acesso Rápido_.
* Clicar na opção _Personalizar Barra de Ferramentas de Acesso Rápido_.

![Image_001](https://github.com/felipebacelo/Suplementos/blob/master/Images/Image_001.png)

Aparecerá a caixa _Opções do Excel_, nesta caixa o usuário deverá selecionar _Macros_ no menu _Escolher comandos em_.

![Image_002](https://github.com/felipebacelo/Suplementos/blob/master/Images/Image_002.png)

Em seguida o usuário poderá adicionar os ícones das respectivas macros a _Barra de Ferramentas de Acesso Rápido_.

![Image_003](https://github.com/felipebacelo/Suplementos/blob/master/Images/Image_003.png)

Após selecionar as macros que deseja utilizar, o usuário deverá configurar o Excel para que o arquivo __Suplementos__ inicialize junto com o Excel.
Para isto é necessário: 
* Clicar em _Guia de Desenvolvedor_.
* Clicar em _Suplementos do Excel_.
* Na caixa de _Suplementos_ o usuário precisará _Procurar_ o diretório em que o arquivo __Suplementos__ está salvo.
* Após localizar o arquivo o usuário deverá marcá-lo e confirmar com OK para que o arquivo __Suplementos__ seja inicializado junto com o Excel.

![Image_004](https://github.com/felipebacelo/Suplementos/blob/master/Images/Image_004.png)
***
### Exemplos de Macros Utilizadas

* Macro utilizada para colorir as células seleciondas com erros em vermelho.
```vba
Sub Erros()
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeConstants, 16).Interior.Color = 255
End Sub
```

* Macro utilizada para converter o valor das células selecionadas em maiúsculas.
```vba
Sub Maiúsculas()
    Dim iCell As Range
        On Error Resume Next
        For Each iCell In Selection
            iCell = UCase(iCell)
        Next iCell
End Sub
```

* Macro utilizada para converter o valor das células selecionadas em minúsculas.
```vba
Sub Minúsculas()
    Dim iCell As Range
        On Error Resume Next
        For Each iCell In Selection
            iCell = LCase(iCell)
        Next iCell
End Sub
```

***
### Licenças

_MIT License_
_Copyright   ©   2020 Felipe Bacelo Rodrigues_
