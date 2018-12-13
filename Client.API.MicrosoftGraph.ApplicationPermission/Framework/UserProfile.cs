using AutoMapper;
using Microsoft.Graph;

namespace Client.API.MicrosoftGraph.ApplicationPermission.Framework
{
    public class UserProfile : Profile
    {
        public UserProfile()
        {
            CreateMap<User, UserModel>()
                .ForMember(dest => dest.Email, opts => opts.MapFrom(src => src.UserPrincipalName));
        }
    }
}
