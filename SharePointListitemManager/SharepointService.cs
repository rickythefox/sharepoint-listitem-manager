using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace SharePointListitemManager
{
    public interface ISharepointService
    {
        IList<string> GetAllSiteUrls();
        IList<string> GetAllListsForSite(string siteUrl);
        SharepointList ExportListToObjects(string siteUrl, string listName);
        void ImportObjectsToList(string siteUrl, string listName, SharepointList list);
        void DeleteAllListItems(string siteUrl, string listName);
    }

    public class SharepointService : ISharepointService
    {
        public IList<string> GetAllSiteUrls()
        {
            var siteNames = new List<string>();
            var currentFarm = SPFarm.Local;
            if (currentFarm == null) return siteNames;

            var service = currentFarm.Services.GetValue<SPWebService>(string.Empty);

            siteNames.AddRange(
                from webApp in service.WebApplications 
                where !webApp.IsAdministrationWebApplication 
                    from site 
                    in webApp.Sites.Cast<SPSite>() 
                    select site.Url
                );
            return siteNames;
        }

        public IList<string> GetAllListsForSite(string siteUrl)
        {
            var web = new SPSite(siteUrl).OpenWeb();
            return (from SPList list in web.Lists select list.Title).ToList();
        }

        public SharepointList ExportListToObjects(string siteUrl, string listName)
        {
            var list = new SPSite(siteUrl).OpenWeb().Lists[listName];
            var cols = new List<SharepointList.Column>();
            
            foreach (SPField field in list.Fields)
            {
                if(field.ReadOnlyField) continue;
                cols.Add(new SharepointList.Column()
                         {
                             Name = field.InternalName,
                             Type = field.FieldValueType
                         });
            }

            var rows = new List<SharepointList.Row>();
            foreach (SPListItem item in list.Items)
            {
                var values = new List<object>();
                foreach (SPField field in list.Fields)
                {
                    if (field.ReadOnlyField) continue;
                    values.Add(item[field.InternalName]);
                }
                rows.Add(new SharepointList.Row() { Values = values});
            }

            return new SharepointList() {Columns = cols, Rows = rows};
        }

        public void ImportObjectsToList(string siteUrl, string listName, SharepointList list)
        {
            var site = new SPSite(siteUrl);
            site.AllowUnsafeUpdates = true;
            var spList = site.OpenWeb().Lists[listName];
            var listItem = spList.Items.Add();

            foreach (var row in list.Rows)
            {
                var col = 0;
                foreach (var value in row.Values)
                {
                    try
                    {
                        listItem[list.Columns[col].Name] = value;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                    col++;
                }
            }
            listItem.Update();
        }

        public void DeleteAllListItems(string siteUrl, string listName)
        {
            var site = new SPSite(siteUrl);
            site.AllowUnsafeUpdates = true;
            var spList = site.OpenWeb().Lists[listName];

            for (var i = spList.ItemCount - 1; i > -1; i--)
            {
                spList.Items[i].Delete();
            }
            spList.Update();
        }
    }
}