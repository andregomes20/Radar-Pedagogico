# 📡 Radar Pedagógico

**Análise diagnóstica e monitoramento educacional offline integrado à Sala do Futuro.**

O **Radar Pedagógico** é um aplicativo desktop desenvolvido em Python focado em centralizar e simplificar a visualização do desempenho escolar de turmas de ensino médio. O sistema opera de forma **100% offline**, garantindo total privacidade e conformidade no manejo de dados educacionais locais.

Sua principal proposta de valor é a automação do fechamento bimestral através da **importação direta da planilha de fechamento da Sala do Futuro**, transformando os dados brutos da secretaria em um painel visual imediato com o panorama completo da sala e dos alunos.

## ⚡ Diferenciais & Proposta Offline

- **Integração com a Sala do Futuro:** Elimina o trabalho manual ao importar diretamente os dados oficiais compilados por bimestre.
- **Privacidade Total (100% Offline):** Nenhum dado de estudantes ou notas sai da máquina local. 
- **Foco na Tomada de Decisão:** Interface desenhada para que a coordenação e os professores visualizem os dados consolidados da sala instantaneamente.

## 🚀 Principais Funcionalidades

- **Importação Bimestral:** Módulo automatizado para leitura e tratamento dos arquivos de fechamento originais da Sala do Futuro.
- **Visualização Consolidada da Sala:** Painel geral que exibe médias, taxas de engajamento e a distribuição de notas da turma selecionada.
- **Filtros por Estudante e Série:** Mecanismo rápido de busca e filtragem para reuniões de conselho de classe ou atendimento pedagógico.

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** [Python](https://python.org)
- **Interface Gráfica:** VS Code / Tkinter *(Ajuste conforme a biblioteca visual exata)*
- **Processamento de Planilhas:** [Pandas](https://pydata.org) (Engine dedicada à leitura e tratamento dos dados da Sala do Futuro)
- **Geração de Documentos:** Bibliotecas nativas ou locais para saída de relatórios offline.

## 💻 Como Executar o Aplicativo

O projeto oferece duas formas de execução: uma simplificada para usuários finais (professores/coordenação) e outra para desenvolvedores.

### Opção 1: Utilizando o Executável (Usuário Final)
1. Acesse a seção de **Releases** do repositório (ou a pasta especificada de distribuição).
2. Baixe o arquivo `RadarPedagogico.exe`.
3. Dê um duplo clique no arquivo para iniciar o aplicativo instantaneamente.
*Nota: Como o app roda 100% offline e processa dados locais, o Windows Defender pode exibir um alerta na primeira execução. Basta clicar em "Mais informações" e "Executar assim mesmo".*

### Opção 2: Rodando via Código Fonte (Desenvolvedor)
Se deseja modificar o código ou rodar o script original, certifique-se de ter o Python instalado:
```bash
# Clone o repositório
git clone https://github.com
cd radar-pedagogico

# Instale as dependências
pip install -r requirements.txt

# Execute a aplicação
python main.py
```

---

## 📂 Fluxo de Trabalho Prático

1. Baixe a **planilha de fechamento da Sala do Futuro** referente ao bimestre atual.
2. Abra o **Radar Pedagógico** (via executável ou script).
3. Utilize o botão de importação para carregar o arquivo local.
4. Escolha o **bimestre** correspondente para processar.
5. Navegue pelo painel para analisar os dados consolidados da sala ou aplicar filtros individuais por aluno.

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** [Python](https://python.org)
- **Processamento de Planilhas:** [Pandas](https://pydata.org)
- **Compilação do Executável:** PyInstaller / Cx_Freeze *(Ajuste se usou outra ferramenta)*

### Notas de Segurança do Build:
- **`--noconsole`**: Oculta o prompt de comando em segundo plano, o que reduz o comportamento suspeito mapeado por heurísticas de antivírus.
- **Isolamento de Bibliotecas**: Garante que dependências de rede sem uso não sejam empacotadas no binário final, eliminando chamadas de conexões externas que acionariam regras restritivas de portas do Windows.

## 🧑‍💻 Autor
- **Prof. André Eduardo Gomes** - *Desenvolvedor & Educador*

## 📄 Licença

Este projeto está sob a licença [MIT](https://choosealicense.com).
Use o código com cuidado.


- 
- **Emissão de Relatórios:** Geração local de relatórios limpos e objetivos formatados para impressão ou salvamento interno.
