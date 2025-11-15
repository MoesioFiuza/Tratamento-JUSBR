using OfficeOpenXml;
using JUSBR_TRAT.Models;

namespace JUSBR_TRAT.Services;

public class PlanilhaService
{
    public async Task<byte[]> ProcessarPlanilhas(
        Stream capaSimplesStream, 
        Stream extraMovimentacoesStream)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Ler CapaSimples
        var capaSimples = LerCapaSimples(capaSimplesStream);

        // Ler Extra_Movimentações (duas abas)
        var dadosProcessos = LerDadosProcessos(extraMovimentacoesStream);
        var movimentacoes = LerMovimentacoes(extraMovimentacoesStream);

        // Processar e gerar planilha tratada
        return await GerarPlanilhaTratada(capaSimples, dadosProcessos, movimentacoes);
    }

    private List<CapaSimples> LerCapaSimples(Stream stream)
    {
        var lista = new List<CapaSimples>();
        
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets[0];
        var rowCount = worksheet.Dimension?.Rows ?? 0;

        // Pular cabeçalho (linha 1)
        for (int row = 2; row <= rowCount; row++)
        {
            if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text))
                break;

            lista.Add(new CapaSimples
            {
                Processo = worksheet.Cells[row, 1].Text,
                DataDistribuicao = worksheet.Cells[row, 2].Text,
                UltimaMovimentacao = worksheet.Cells[row, 3].Text,
                ClasseJudicial = worksheet.Cells[row, 4].Text,
                Assunto = worksheet.Cells[row, 5].Text,
                Partes = worksheet.Cells[row, 6].Text,
                OrgaoJulgador = worksheet.Cells[row, 7].Text
            });
        }

        return lista;
    }

    private List<DadosProcesso> LerDadosProcessos(Stream stream)
    {
        var lista = new List<DadosProcesso>();
        
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Dados"];
        
        if (worksheet == null)
            return lista;

        var rowCount = worksheet.Dimension?.Rows ?? 0;

        // Pular cabeçalho (linha 1)
        for (int row = 2; row <= rowCount; row++)
        {
            if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text))
                break;

            lista.Add(new DadosProcesso
            {
                Processo = worksheet.Cells[row, 1].Text,
                StatusRaspagem = worksheet.Cells[row, 2].Text,
                OrgaoJulgador = worksheet.Cells[row, 3].Text,
                AutuadoEm = worksheet.Cells[row, 4].Text,
                ValorCausa = worksheet.Cells[row, 5].Text,
                SegredoSigilo = worksheet.Cells[row, 6].Text,
                ClasseJudicial = worksheet.Cells[row, 7].Text,
                PoloAtivo = worksheet.Cells[row, 8].Text,
                PoloPassivo = worksheet.Cells[row, 9].Text,
                UltimaMovimentacao = worksheet.Cells[row, 10].Text
            });
        }

        return lista;
    }

    private List<Movimentacao> LerMovimentacoes(Stream stream)
    {
        var lista = new List<Movimentacao>();
        
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets["Movimentações"];
        
        if (worksheet == null)
            return lista;

        var rowCount = worksheet.Dimension?.Rows ?? 0;

        // Pular cabeçalho (linha 1)
        for (int row = 2; row <= rowCount; row++)
        {
            if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text))
                break;

            lista.Add(new Movimentacao
            {
                Processo = worksheet.Cells[row, 1].Text,
                Data = worksheet.Cells[row, 2].Text,
                Descricao = worksheet.Cells[row, 3].Text,
                Ordem = worksheet.Cells[row, 4].Text
            });
        }

        return lista;
    }

    // Métodos auxiliares para extrair informações do Órgão Julgador
    private string ExtrairComarca(string orgaoJulgador)
    {
        if (string.IsNullOrWhiteSpace(orgaoJulgador))
            return "";

        // Procurar por padrão: número seguido de "ª" ou "º" ou espaço/letra
        // Exemplos: "5ª", "16VARA", "43 CIVEL"
        var match = System.Text.RegularExpressions.Regex.Match(orgaoJulgador, @"-?\s*(\d+)[ªº]?\s*[A-Z]");
        if (match.Success)
        {
            return match.Groups[1].Value;
        }

        // Tentar padrão alternativo: número no início após hífen
        match = System.Text.RegularExpressions.Regex.Match(orgaoJulgador, @"-\s*(\d+)");
        if (match.Success)
        {
            return match.Groups[1].Value;
        }

        return "";
    }

    private string ExtrairEstado(string orgaoJulgador)
    {
        if (string.IsNullOrWhiteSpace(orgaoJulgador))
            return "";

        // Procurar por padrão TJ seguido de duas letras
        var match = System.Text.RegularExpressions.Regex.Match(orgaoJulgador, @"TJ([A-Z]{2})");
        if (match.Success)
        {
            return match.Groups[1].Value;
        }

        // Se for TRF, não tem estado
        return "";
    }

    private string ExtrairOrgao(string orgaoJulgador)
    {
        if (string.IsNullOrWhiteSpace(orgaoJulgador))
            return "";

        // Primeiro, pegar a parte antes do "-"
        string parteAntesHifen = "";
        var index = orgaoJulgador.IndexOf(" - ");
        if (index > 0)
        {
            parteAntesHifen = orgaoJulgador.Substring(0, index).Trim();
        }
        else
        {
            index = orgaoJulgador.IndexOf("-");
            if (index > 0)
            {
                parteAntesHifen = orgaoJulgador.Substring(0, index).Trim();
            }
            else
            {
                parteAntesHifen = orgaoJulgador.Trim();
            }
        }

        // Se começar com "TJ" seguido de letras (ex: TJRN, TJSP, TJAM), separar como "TJ-RN", "TJ-SP", "TJ-AM"
        if (parteAntesHifen.StartsWith("TJ") && parteAntesHifen.Length > 2)
        {
            var match = System.Text.RegularExpressions.Regex.Match(parteAntesHifen, @"^TJ([A-Z]{2,})$");
            if (match.Success)
            {
                return "TJ-" + match.Groups[1].Value;
            }
        }

        // Se for "TRF1" ou começar com "TRF", manter como está (sem adicionar hífen)
        if (parteAntesHifen.StartsWith("TRF"))
        {
            return parteAntesHifen;
        }

        // Para outros casos, retornar como está
        return parteAntesHifen;
    }

    private string ExtrairComarcaAposNumeral(string orgaoJulgador)
    {
        if (string.IsNullOrWhiteSpace(orgaoJulgador))
            return "";

        // Procurar por padrão: número seguido de "ª" ou "º" ou espaço/letra
        // Exemplos: "5ª VARA CÍVEL DA CAPITAL" → "VARA CÍVEL DA CAPITAL"
        // "16VARA DO SISTEMA..." → "VARA DO SISTEMA..."
        // "43 CIVEL DE CENTRAL" → "CIVEL DE CENTRAL"
        var match = System.Text.RegularExpressions.Regex.Match(orgaoJulgador, @"\d+[ªº]?\s*([A-Z].*)");
        if (match.Success)
        {
            return match.Groups[1].Value.Trim();
        }

        // Tentar padrão alternativo: número seguido de espaço
        match = System.Text.RegularExpressions.Regex.Match(orgaoJulgador, @"\d+\s+(.+)");
        if (match.Success)
        {
            return match.Groups[1].Value.Trim();
        }

        return "";
    }

    private async Task<byte[]> GerarPlanilhaTratada(
        List<CapaSimples> capaSimples, 
        List<DadosProcesso> dadosProcessos,
        List<Movimentacao> movimentacoes)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("PlanilhaTratada");

        var dataAtual = DateTime.Now.ToString("dd/MM/yyyy");

        // Criar dicionários para busca rápida
        var dadosPorProcesso = dadosProcessos.ToDictionary(d => d.Processo ?? "");
        var movimentacoesPorProcesso = movimentacoes
            .GroupBy(m => m.Processo ?? "")
            .ToDictionary(g => g.Key, g => g.OrderByDescending(m => m.Data).ToList());

        // Definir cabeçalhos conforme a ordem especificada
        var cabecalhos = new[]
        {
            "pasta", "numeroProcessoAnterior", "cnj", "tipoPartePoloAtivo", "partePoloAtivo",
            "tipoPartePoloPassivo", "partePoloPassivo", "cliente", "tipoDeRito", "dataDistribuicao",
            "numeroUnidade", "unidade", "especialidade", "comarca", "estado", "orgao", "natureza",
            "materia", "dataInstancia", "tipoInstancia", "sistemaExterno", "processoEletronico",
            "processoEstrategico", "valorCausa", "valorFinalCausa", "tipoAcao", "tipoObjeto",
            "dataFase", "fase", "dataStatus", "status", "grupoProcesso", "prioridadeDe",
            "data_resultado", "tipo_resultado", "descricao_resultado", "dataEvento", "tipoEvento",
            "descricaoEvento", "complementoEvento", "observacaoEvento", "solicitanteEvento",
            "responsavelEvento", "grupoTrabalho", "corresponsavel", "dataNotificacao",
            "dataNotificacaoAdicional", "probabilidadePerda", "dataValorProvisionado",
            "valorProvisionado", "dataAndamento", "tipoAndamento", "descricaoAndamento",
            "complementoAndamento", "solicitanteAndamento", "responsavelAndamento",
            "corresponsavelAndamento", "descricaoObjeto", "escritorioCredenciado",
            "dataContratacao", "observacaoDoProcesso", "parecerDoProcesso"
        };

        // Escrever cabeçalhos
        for (int col = 1; col <= cabecalhos.Length; col++)
        {
            worksheet.Cells[1, col].Value = cabecalhos[col - 1];
        }

        // Estilizar cabeçalho
        using (var range = worksheet.Cells[1, 1, 1, cabecalhos.Length])
        {
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
        }

        int row = 2;
        foreach (var capa in capaSimples)
        {
            var processo = capa.Processo ?? "";
            dadosPorProcesso.TryGetValue(processo, out var dados);
            movimentacoesPorProcesso.TryGetValue(processo, out var movs);
            var ultimaMov = movs?.FirstOrDefault();

            int col = 1;

            // pasta - Deixe vazio
            worksheet.Cells[row, col++].Value = "";

            // numeroProcessoAnterior - Deixe vazio
            worksheet.Cells[row, col++].Value = "";

            // cnj - Coluna 'Processo' da planilha Capa
            worksheet.Cells[row, col++].Value = capa.Processo;

            // tipoPartePoloAtivo - "Autor"
            worksheet.Cells[row, col++].Value = "Autor";

            // partePoloAtivo - Coluna "Polo Ativo" da planilha Extra 'aba' Dados
            worksheet.Cells[row, col++].Value = dados?.PoloAtivo ?? "";

            // tipoPartePoloPassivo - "Réu"
            worksheet.Cells[row, col++].Value = "Réu";

            // partePoloPassivo - Coluna "Polo Passivo" da planilha Extra 'aba' Dados
            worksheet.Cells[row, col++].Value = dados?.PoloPassivo ?? "";

            // cliente - Coluna "Polo Passivo" da planilha Extra 'aba' Dados
            worksheet.Cells[row, col++].Value = dados?.PoloPassivo ?? "";

            // tipoDeRito - "Sumarissimo"
            worksheet.Cells[row, col++].Value = "Sumarissimo";

            // dataDistribuicao - Coluna "Autuado em" da planilha Extra 'aba' Dados
            worksheet.Cells[row, col++].Value = dados?.AutuadoEm ?? "";

            // numeroUnidade - Extrair número do Órgão Julgador da planilha Extra 'aba' Dados
            var orgaoJulgador = dados?.OrgaoJulgador ?? "";
            worksheet.Cells[row, col++].Value = ExtrairComarca(orgaoJulgador);

            // unidade - "Juizado Especial"
            worksheet.Cells[row, col++].Value = "Juizado Especial";

            // especialidade - "Cível"
            worksheet.Cells[row, col++].Value = "Cível";

            // comarca - Tudo que vem após o numeral do Órgão Julgador da planilha Extra 'aba' Dados
            worksheet.Cells[row, col++].Value = ExtrairComarcaAposNumeral(orgaoJulgador);

            // estado - Extrair duas letras após TJ do Órgão Julgador da planilha Extra 'aba' Dados
            worksheet.Cells[row, col++].Value = ExtrairEstado(orgaoJulgador);

            // orgao - Parte antes do "-" do Órgão Julgador da planilha Extra 'aba' Dados
            worksheet.Cells[row, col++].Value = ExtrairOrgao(orgaoJulgador);

            // natureza - "Judicial"
            worksheet.Cells[row, col++].Value = "Judicial";

            // materia - "Cível"
            worksheet.Cells[row, col++].Value = "Cível";

            // dataInstancia - Mesma da dataDistribuicao (Coluna "Autuado em" da planilha Extra 'aba' Dados)
            worksheet.Cells[row, col++].Value = dados?.AutuadoEm ?? "";

            // tipoInstancia - "1ª Instância"
            worksheet.Cells[row, col++].Value = "1ª Instância";

            // sistemaExterno - Vazio
            worksheet.Cells[row, col++].Value = "";

            // processoEletronico - "Sim"
            worksheet.Cells[row, col++].Value = "Sim";

            // processoEstrategico - Vazio
            worksheet.Cells[row, col++].Value = "";

            // valorCausa - Coluna "Valor da Causa" da planilha Extra 'aba' Dados
            worksheet.Cells[row, col++].Value = dados?.ValorCausa ?? "";

            // valorFinalCausa - Vazio
            worksheet.Cells[row, col++].Value = "";

            // tipoAcao - "Reclamação do Consumidor"
            worksheet.Cells[row, col++].Value = "Reclamação do Consumidor";

            // tipoObjeto - Vazio
            worksheet.Cells[row, col++].Value = "";

            // dataFase - Vazio
            worksheet.Cells[row, col++].Value = "";

            // fase - Vazio
            worksheet.Cells[row, col++].Value = "";

            // dataStatus - Data da execução do programa em dd/mm/aaaa
            worksheet.Cells[row, col++].Value = dataAtual;

            // status - "Ativo"
            worksheet.Cells[row, col++].Value = "Ativo";

            // grupoProcesso - Vazio
            worksheet.Cells[row, col++].Value = "";

            // prioridadeDe - Vazio
            worksheet.Cells[row, col++].Value = "";

            // data_resultado - Vazio
            worksheet.Cells[row, col++].Value = "";

            // tipo_resultado - Vazio
            worksheet.Cells[row, col++].Value = "";

            // descricao_resultado - Vazio
            worksheet.Cells[row, col++].Value = "";

            // dataEvento - "1035"
            worksheet.Cells[row, col++].Value = "1035";

            // tipoEvento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // descricaoEvento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // complementoEvento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // observacaoEvento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // solicitanteEvento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // responsavelEvento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // grupoTrabalho - Vazio
            worksheet.Cells[row, col++].Value = "";

            // corresponsavel - "32"
            worksheet.Cells[row, col++].Value = "32";

            // dataNotificacao - "Agora" (data atual)
            worksheet.Cells[row, col++].Value = dataAtual;

            // dataNotificacaoAdicional - "No dia do evento" (data atual)
            worksheet.Cells[row, col++].Value = dataAtual;

            // probabilidadePerda - "Possível"
            worksheet.Cells[row, col++].Value = "Possível";

            // dataValorProvisionado - Data da execução do programa em dd/mm/aaaa
            worksheet.Cells[row, col++].Value = dataAtual;

            // valorProvisionado - Vazio
            worksheet.Cells[row, col++].Value = "";

            // dataAndamento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // tipoAndamento - "Não informado"
            worksheet.Cells[row, col++].Value = "Não informado";

            // descricaoAndamento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // complementoAndamento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // solicitanteAndamento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // responsavelAndamento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // corresponsavelAndamento - Vazio
            worksheet.Cells[row, col++].Value = "";

            // descricaoObjeto - "Não Informado"
            worksheet.Cells[row, col++].Value = "Não Informado";

            // escritorioCredenciado - "VALENÇA & ASSOCIADOS"
            worksheet.Cells[row, col++].Value = "VALENÇA & ASSOCIADOS";

            // dataContratacao - Vazio
            worksheet.Cells[row, col++].Value = "";

            // observacaoDoProcesso - Vazio
            worksheet.Cells[row, col++].Value = "";

            // parecerDoProcesso - Vazio
            worksheet.Cells[row, col++].Value = "";

            row++;
        }

        // Ajustar largura das colunas
        worksheet.Cells.AutoFitColumns();

        return await Task.FromResult(package.GetAsByteArray());
    }
}