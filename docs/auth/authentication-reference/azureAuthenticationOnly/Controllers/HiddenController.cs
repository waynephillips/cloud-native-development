using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace azureAuthenticationOnly.Controllers
{
    public class HiddenController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}