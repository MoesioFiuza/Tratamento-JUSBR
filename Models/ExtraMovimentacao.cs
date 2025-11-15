namespace JUSBR_TRAT.Models;

public class DadosProcesso
{
    public string? Processo { get; set; }
    public string? StatusRaspagem { get; set; }
    public string? OrgaoJulgador { get; set; }
    public string? AutuadoEm { get; set; }
    public string? ValorCausa { get; set; }
    public string? SegredoSigilo { get; set; }
    public string? ClasseJudicial { get; set; }
    public string? PoloAtivo { get; set; }
    public string? PoloPassivo { get; set; }
    public string? UltimaMovimentacao { get; set; }
}

public class Movimentacao
{
    public string? Processo { get; set; }
    public string? Data { get; set; }
    public string? Descricao { get; set; }
    public string? Ordem { get; set; }
}