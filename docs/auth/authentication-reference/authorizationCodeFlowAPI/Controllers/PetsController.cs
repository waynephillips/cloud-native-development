using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;

namespace authorizationCodeFlowAPI.Controllers
{
    [Route("api/[controller]")]
    public class PetsController : Controller
    {
        [HttpGet]
        public ActionResult Get()
        {
            return Ok(new List<string>() { "cat", "dog", "bird"});
        }

        [Authorize]
        [HttpGet("{number}")]
        public ActionResult Get(int number)
        {
            return Ok(new List<string>() { "cow", "duck", "koala"});
        }
    }
}