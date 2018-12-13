#region using
using AutoMapper;
using Client.API.MicrosoftGraph.ApplicationPermission.Framework;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Threading.Tasks;
#endregion

namespace Client.API.MicrosoftGraph.ApplicationPermission.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UserController : ControllerBase
    {
        private readonly MsGraphService _msGraphService;
        private readonly IMapper _mapper;

        public UserController(MsGraphService msGraphService, IMapper mapper)
        {
            _msGraphService = msGraphService;
            _mapper = mapper;
        }

        /// <summary>
        /// Retrieve all users
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<IList<UserModel>> Get()
        {
            var client = _msGraphService.GetAuthenticatedClient();
            var result = await client.Users.Request().GetAsync();

            return _mapper.Map<IList<User>, IList<UserModel>>(result.CurrentPage);
        }

        /// <summary>
        /// Retrieve user by id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet("{id}")]
        public async Task<ActionResult<UserModel>> Get(string id)
        {
            var client = _msGraphService.GetAuthenticatedClient();
            var result = await client.Users[id].Request().GetAsync();

            return _mapper.Map<User, UserModel>(result);
        }
    }
}