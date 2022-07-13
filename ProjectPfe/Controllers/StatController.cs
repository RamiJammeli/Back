using ConnexionMongo.Services;
using Microsoft.AspNetCore.Mvc;
using ProjectPfe.Models;
using ProjectPfe.Models.stat;
using ProjectPfe.Services;

namespace ProjectPfe.Controllers
{
    [Produces("application/json")]
    [ApiController]
    [Route("[Controller]")]
    public class StatController : ControllerBase
    {

        private readonly CategorieService categorieService;
        private readonly IntegrationService integrationService;
        public StatController(CategorieService _categorieService, IntegrationService _integrationService)
        {
            categorieService = _categorieService;
            integrationService = _integrationService;
        }

        [HttpGet(Name = "BestCategorie")]
        [Route("BestCategorie")]
        public List<StatModel> BestCategorie()
        {
            var integrations = integrationService.GetIntegrationGroup();
           return integrations.GroupBy(x => x.categorie.libelle).Select(x => new StatModel { Libelle= x.Key, Nombre=x.Count()  } ).ToList();
            

        }


        [HttpGet(Name = "BestUsers")]
        [Route("BestUsers")]
        public List<StatModel> BestUsers()
        {
            var integrations = integrationService.GetIntegrationGroup();
            return integrations.GroupBy(x => x.UserImport.Username).Select(x => new StatModel { Libelle = x.Key, Nombre = x.Count() }).ToList();


        }


    }
}
