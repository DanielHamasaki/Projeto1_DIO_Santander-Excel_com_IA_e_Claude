# 📊 Tabela e Análise de Dados — Excel

Projeto desenvolvido como parte do desafio **Tabela e Análise de Dados** do curso **Santander — Excel com IA e Claude**, da DIO.

O arquivo apresenta uma planilha de apoio ao planejamento de investimentos, utilizando recursos do Excel para entrada de dados, cálculos financeiros, cenários de longo prazo e distribuição de investimentos em FIIs conforme um perfil selecionado.

## 📁 Estrutura do projeto

O arquivo `Projeto1_DIO_Santander-Excel_com_IA_e_Claude.xlsx` possui duas planilhas:

### `APP`

É a planilha principal do projeto. Nela estão concentradas as configurações, os cálculos de investimento, os cenários e a distribuição mensal sugerida entre tipos de FIIs.

### `Tab_apoio`

É a tabela de apoio utilizada para relacionar o perfil selecionado aos percentuais sugeridos para cada tipo de FII.

A tabela contém três perfis:

- Conservador
- Moderado
- Agressivo

E seis tipos de FIIs:

- PAPEL
- TIJOLO
- HÍBRIDOS
- FOFs
- DESENVOLVIMENTO
- HOTELARIAS

## 🧮 Recursos do Excel utilizados

O projeto utiliza os seguintes recursos, identificados diretamente no arquivo:

- Fórmulas e referências entre células
- Funções financeiras `FV`
- Funções `VLOOKUP` e `SUM`
- Operações matemáticas com percentuais
- Nomes definidos para células e variáveis
- Validação de dados com lista suspensa
- Tabela de apoio para consulta de percentuais
- Gráfico de pizza
- Células mescladas para organização visual
- Formatação de valores e percentuais

## ⚙️ Configurações

Na seção **CONFIGURAÇÕES**, o usuário informa:

- **Salário**
- **Rendimento da Carteira**

A planilha calcula automaticamente uma **Sugestão de investimento**, correspondente a 30% do salário.

A fórmula utilizada é:

```excel
=salario*30%
```

## 💰 Simulação de investimento mensal

A seção **INVESTIMENTO MENSAL** permite informar:

- Quanto investir ao mês
- Por quantos anos
- Taxa de rendimento mensal

A partir desses valores, a planilha calcula:

- **Patrimônio acumulado**
- **Dividendos mensais**

O patrimônio acumulado utiliza a função financeira `FV`:

```excel
=FV(taxa_mensal,periodo*12,aporte*-1)
```

Os dividendos mensais são calculados a partir do patrimônio acumulado e do rendimento informado:

```excel
=patrimonio_acumulado*rendimento_carteira
```

## 📈 Cenários

A seção **CENÁRIOS** apresenta simulações para:

- 2 anos
- 5 anos
- 10 anos
- 20 anos
- 30 anos

Para cada período, são calculados:

- Patrimônio acumulado
- Dividendos correspondentes

O cálculo do patrimônio utiliza a mesma função `FV`, considerando o aporte mensal e a taxa de rendimento mensal.

## 🏦 Distribuição entre tipos de FIIs

A seção **PERFIL** permite selecionar um dos três perfis disponíveis por meio de uma **lista suspensa**:

- Conservador
- Moderado
- Agressivo

Depois de selecionado o perfil e informado o **Valor Mensal**, a planilha consulta a tabela `Tab_apoio` para obter o percentual correspondente a cada tipo de FII.

A consulta é realizada utilizando `VLOOKUP`:

```excel
=VLOOKUP($D$29&"-"&B33,Tab_apoio!$B$3:$E$20,4,FALSE)
```

O projeto utiliza uma chave formada pela combinação:

```text
Perfil-Tipo de FII
```

Exemplo:

```text
Conservador-PAPEL
```

Com o percentual encontrado, a planilha calcula o valor mensal destinado a cada categoria:

```excel
=C33*$D$30
```

Ao final, a linha **TOTAL** soma os percentuais e os valores distribuídos:

```excel
=SUM(C33:C38)
```

```excel
=SUM(D33:D38)
```

## 📊 Visualização

A planilha principal contém um **gráfico de pizza** relacionado aos percentuais sugeridos para os tipos de FIIs.

O gráfico utiliza os dados da distribuição percentual apresentada na seção **TIPOS DE FIIs**.

## 🏷️ Nomes definidos

O arquivo utiliza nomes definidos para facilitar a leitura das fórmulas:

| Nome | Referência |
|---|---|
| `salario` | `APP!$D$11` |
| `rendimento_carteira` | `APP!$D$12` |
| `sugestao_invest` | `APP!$D$13` |
| `aporte` | `APP!$D$16` |
| `periodo` | `APP!$D$17` |
| `taxa_mensal` | `APP!$D$18` |
| `patrimonio_acumulado` | `APP!$D$19` |

## 🔄 Fluxo da planilha

```text
Salário
   ↓
Sugestão de investimento
   ↓
Aporte mensal + período + taxa
   ↓
Patrimônio acumulado
   ↓
Dividendos mensais
   ↓
Seleção do perfil
   ↓
Consulta dos percentuais na Tab_apoio
   ↓
Distribuição do valor entre os tipos de FIIs
   ↓
Gráfico de distribuição
```

## 📌 Observação

Este projeto representa uma aplicação prática de recursos do Excel para **organização, cálculo e visualização de dados relacionados a uma simulação de investimentos**. Os valores e percentuais presentes no arquivo são os utilizados na própria planilha e servem como dados da simulação.

---

### 🛠️ Tecnologias e recursos

- Microsoft Excel
- Fórmulas financeiras
- Funções de busca
- Funções de soma
- Validação de dados
- Nomes definidos
- Gráfico de pizza
- Tabela de apoio
