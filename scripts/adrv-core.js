var __HS = __HS || {}; 
__HS.Modules = __HS.Modules || {};

if (console) {
    console.log('Loading Aderant Drive Core library.  Type of __HS is ' + typeof __HS);
}

(function (ns) {

    /*
      Namespace : __HS.Modules.adrv
      Provides useful functions for interfacing Handshake elements rendered by the Aderant Drive web Part with 
      Sharepoint Elements. Also includes methods that are neccessary only in the context of a skin renedered
      on an Aderant Drive site rather than a realtime skin renderning from the HS server.
    */

    ns.isSharePoint = function() {
        // return true if sharepoint.com is any part of the URL 
        return location.href.indexOf('sharepoint.com') > 0; 
    }

    ns.baseSiteUrl = function () {
        /*
        This skin will never be on a root site, only on a sitecollection 
        which will (? should) always be https://<tenant>.sharepoint.com/sites/<sitename>
        */
        var parts = location.href.split('/');
        if (parts.length > 3) {
            return parts[0] + '//' + parts[2] + '/' + parts[3] + '/' + parts[4]
        } else {
            return '';
        }
    }

    ns.openSitePage = function (pageName) {
        /*
        Open the named page from the SitePages library of the current site.
        */
        if (pageName) {
            location.href = ns.baseSiteUrl() + '/SitePages/' + pageName + '.aspx';
        }
    }

    ns.getHubSiteMembers = function (hubid, startrow, allSites) {
        /*
        Use the Sharepoint Online Search REST API to query for the Sites associated to the Client Hub Site that the current user has permissions to. 
        Note - the {} around the hubid guid are expected and required, so they will be added if omitted in the parameter
        */
        allSites = allSites || [];
        startrow = startrow || 0;
        if (hubid.length === 36) hubid = '{' + hubid + '}';
        var hardLimit = 10000;
        var rowlimit = 500;
        var url = ns.baseSiteUrl() + "/_api/search/query?querytext='DepartmentId:" + hubid + " contentclass:STS_Site'&selectproperties='Title,Path,DepartmentId'&trimduplicates=false&RowLimit=" + rowlimit + "&startrow=" + startrow;

        return jQuery.getJSON(url).then(function (data) {
            var relevantResults = data.PrimaryQueryResult.RelevantResults;
            allSites = allSites.concat(relevantResults.Table.Rows);

            // run recursively until we get all available results back.
            if (relevantResults.TotalRows > startrow + relevantResults.RowCount & allSites.length < hardLimit) {
                return ns.getHubSiteMembers(hubid, startrow + relevantResults.RowCount, allSites);
            }

            return allSites;
        });
    }

    ns.getSiteInfo = function () {
        return jQuery.getJSON(ns.baseSiteUrl() + '/_api/site');
    }

    ns.onInitSecurityTrimMatter = function(options) {
        if (ns.isSharePoint()) {
            options.autoBind = false; 
        }
    }

    ns.matterSecurityTrim = function (myMatterControl) {
        /*
        matterSecurityTrim is used to filter out matters from an HTML5List or Grid to which the current user does NOT have permissions.
        Assumptions / Setup: 
        1. myMatterControl a jQuery selector for either the Kendo Grid or Kendo List View 
        2. The list or grid has a field named 'sitecollectionurl' which must be the same as the url returned in the Path field from SP Search
        3. The list of grid should have an oninitialize property set to  "AderantDrive.onInitSecurityTrimMatter(options);"
        4. The skin that contains the list/grid should have the following code in the body of the skin
            <script>
                AderantDrive.matterSecurityTrim('[hsname="matter-grid"]'); 
            </script>
        note: this script cannot be made part of the grid itself. it should be called independently. 
              if the skin  is being rendered anywhere other than a SharePoint Online Page, it does nothint. 
        */
        if (ns.isSharePoint()) {
            var myMatterSites = [];
            var myMatterSitesFilter = {
                logic: "and",
                filters: [{
                    field: "sitecollectionurl",
                    operator: function (item) {
                        // item is the value of the sitecollectionurl field 
                        // return true if this item is in the list of matter sites 
                        return (myMatterSites.indexOf(item) !== -1)
                    },
                    value: ''
                }]
            };

            function addSiteFilter() {
                // Private function for ns.matterSecurityTrim  
                var role, controlName;
                var elem = jQuery(myMatterControl)[0];
                if (elem) {
                    role = jQuery(elem).data('role');
                    controlName = (role == 'grid') ? "kendoGrid" : "kendoListView";
                    if (controlName) {
                        var control = jQuery(elem).data(controlName);
                    }
                }
                if (elem && control && control.dataSource) {
                    control.dataSource.filter(myMatterSitesFilter)
                } else {
                    if (console) console.log('No ' + controlName + ' Object found, trying again...');
                    setTimeout(addSiteFilter, 250);
                }
            }

            var siteInfo = ns.getSiteInfo();
            siteInfo.done(function (siteinfo) {
                if (siteinfo && siteinfo.HubSiteId) {
                    ns.getHubSiteMembers(siteinfo.HubSiteId, 0)
                        .done(function (results) {
                            results.forEach(function (Rows) {
                                Rows.Cells.forEach(function (item) {
                                    if (item.Key === 'Path') {
                                        myMatterSites.push(item.Value);
                                    }
                                });
                            });
                            addSiteFilter();
                        });
                } else {
                    if (console) console.log('Error - did not security trim matter list...');
                }
            });
        }
    }
})(__HS.Modules.adrv = __HS.Modules.adrv || {});
