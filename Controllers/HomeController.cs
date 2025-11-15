using Microsoft.AspNetCore.Mvc;
using JUSBR_TRAT.Models;
using JUSBR_TRAT.Services;

namespace JUSBR_TRAT.Controllers;

public class HomeController : Controller
{
    private readonly PlanilhaService _planilhaService;
    private readonly ILogger<HomeController> _logger;

    public HomeController(PlanilhaService planilhaService, ILogger<HomeController> logger)
    {
        _planilhaService = planilhaService;
        _logger = logger;
    }

    [HttpGet]
    public IActionResult Index()
    {
        // Retornar o arquivo HTML estático do wwwroot
        return PhysicalFile(
            Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "index.html"),
            "text/html");
    }

    [HttpPost]
    [DisableRequestSizeLimit]
    [RequestFormLimits(MultipartBodyLengthLimit = 104857600)] // 100MB
    public async Task<IActionResult> Processar([FromForm] ProcessamentoViewModel model)
    {
        _logger.LogInformation("Processar chamado. Model: {ModelNull}, CapaSimplesFile: {CapaNull}, ExtraMovimentacoesFile: {ExtraNull}", 
            model == null,
            model?.CapaSimplesFile == null, 
            model?.ExtraMovimentacoesFile == null);

        if (model == null)
        {
            _logger.LogWarning("Model é null");
            return BadRequest(new { error = "Dados inválidos." });
        }

        if (model.CapaSimplesFile == null || model.ExtraMovimentacoesFile == null)
        {
            _logger.LogWarning("Arquivos não foram enviados. CapaSimplesFile: {CapaNull}, ExtraMovimentacoesFile: {ExtraNull}", 
                model.CapaSimplesFile == null, 
                model.ExtraMovimentacoesFile == null);
            return BadRequest(new { error = "Por favor, envie ambas as planilhas." });
        }

        if (model.CapaSimplesFile.Length == 0 || model.ExtraMovimentacoesFile.Length == 0)
        {
            _logger.LogWarning("Arquivos estão vazios");
            return BadRequest(new { error = "Os arquivos enviados estão vazios." });
        }

        try
        {
            _logger.LogInformation("Iniciando processamento das planilhas. CapaSimples: {CapaSize} bytes, ExtraMovimentacoes: {ExtraSize} bytes", 
                model.CapaSimplesFile.Length, 
                model.ExtraMovimentacoesFile.Length);
            
            using var capaStream = model.CapaSimplesFile.OpenReadStream();
            using var movStream = model.ExtraMovimentacoesFile.OpenReadStream();

            var planilhaTratada = await _planilhaService.ProcessarPlanilhas(capaStream, movStream);

            _logger.LogInformation("Planilha processada com sucesso. Tamanho: {Tamanho} bytes", planilhaTratada.Length);

            if (planilhaTratada == null || planilhaTratada.Length == 0)
            {
                _logger.LogError("Planilha tratada está vazia");
                return BadRequest(new { error = "Erro ao gerar planilha tratada." });
            }

            return File(planilhaTratada, 
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                $"PlanilhaTratada_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro ao processar planilhas: {Mensagem}", ex.Message);
            return StatusCode(500, new { error = $"Erro ao processar: {ex.Message}" });
        }
    }
}