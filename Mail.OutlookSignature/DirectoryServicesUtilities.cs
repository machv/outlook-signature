﻿using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace Mail.OutlookSignature
{
    public class DirectoryServicesUtilities
    {
        public static IEnumerable<string> GetGroupsOfUser(string samAccountName)
        {
            var userNestedMembership = new List<string>();

            var domainConnection = new DirectoryEntry();
            domainConnection.AuthenticationType = AuthenticationTypes.Secure;

            var samSearcher = new DirectorySearcher();

            samSearcher.SearchRoot = domainConnection;
            samSearcher.Filter = "(samAccountName=" + samAccountName + ")";
            samSearcher.PropertiesToLoad.Add("displayName");

            var samResult = samSearcher.FindOne();

            if (samResult != null)
            {
                var theUser = samResult.GetDirectoryEntry();
                theUser.RefreshCache(new string[] { "tokenGroups" });

                foreach (byte[] resultBytes in theUser.Properties["tokenGroups"])
                {
                    var SID = new SecurityIdentifier(resultBytes, 0);
                    var sidSearcher = new DirectorySearcher();

                    sidSearcher.SearchRoot = domainConnection;
                    sidSearcher.Filter = "(objectSid=" + SID.Value + ")";
                    sidSearcher.PropertiesToLoad.Add("name");

                    var sidResult = sidSearcher.FindOne();
                    if (sidResult != null)
                    {
                        userNestedMembership.Add((string)sidResult.Properties["name"][0]);
                    }
                }
            }

            return userNestedMembership;
        }
    }
}
