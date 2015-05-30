window.config = {
    tenant: '',
    clientId: '',
    postLogoutRedirectUri: window.location.origin,
    cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost.
};

var authContext;
var token;

Polymer({
    userName: '',
    tenant: '',
    clientID: '',
    token: '',
    ordination: '',
    refinement: '',
    refinementFilters: '',
    sortingList:  '',
    resultElement: '',
    totalElement: '',
    pageElement: '10',
    ready: function () {
        window.config.tenant = this.tenant;
        window.config.clientId = this.clientID;

        authContext = new AuthenticationContext(config);
        // Acquire token for Files resource.

        authContext.acquireToken("https://" + this.tenant + "-my.sharepoint.com", function (error, accessToken) {

            // Handle ADAL Errors.
            if (error || !accessToken) {
                console.log('ADAL error occurred: ' + error);
                return;
            }

            token = accessToken;

        });
    },
    attached: function()
    {
        var searchBox = document.createElement("input");
        searchBox.type = "text";
        searchBox.className = "searchbox-input";
        searchBox.id = "searchbox";

        var button = document.createElement("button");
        button.innerText = "Search";
        button.className = "search-button"
        button.addEventListener("click", this.search.bind(this), true);

        if (this.ordination != '')
            this.loadOrdination();

        this.appendChild(searchBox);
        this.appendChild(button);
    },
    search: function()
    {
        var searchText = document.getElementById("searchbox").value;

        if (searchText.length != 0)
        {
            var request = new XMLHttpRequest();

            var requestText = "https://" + this.tenant + "-my.sharepoint.com/_api/search/query?querytext='" + searchText + "'&rowlimit=" + this.pageElement;

            this.refinementFilters = '';
            this.sortingList = '';

            if (this.refinement != '')
                requestText += "&refiners='" + this.refinement + "'";

            request.open("GET", requestText, true);
            request.setRequestHeader("Authorization", "Bearer " + token);
            request.setRequestHeader("Access-Control-Allow-Origin", "*");
            request.setRequestHeader("accept", "application/json;odata=verbose");
            request.setRequestHeader("content-type", "application/json;odata=verbose");

            request.parent = this

            request.onreadystatechange = function () {
                if (request.readyState == 4 && request.status == 200) {
                    var jsonResponseSearch = JSON.parse(request.responseText);

                    this.parent.totalElement = jsonResponseSearch.d.query.PrimaryQueryResult.RelevantResults.TotalRows;

                    if (this.parent.refinement != '')
                        this.parent.loadRefiner(jsonResponseSearch.d.query.PrimaryQueryResult.RefinementResults.Refiners.results);

                    this.parent.showResults(jsonResponseSearch.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results, 1);
                }
            }

            request.send();
        }
    },
    loadRefiner : function(refiners)
    {
        var listToRemoved = document.getElementById("refiner-list");

        if (listToRemoved != null)
            this.removeChild(listToRemoved);

        var refinerList = document.createElement("ul");
        refinerList.id = "refiner-list";

        for(var i = 0; i < refiners.length; i++)
        {
            var refiner = document.createElement("li");
            
            var spanTitle = document.createElement("span");
            spanTitle.innerText = refiners[i].Name;

            var content = document.createElement("ul");

            results = refiners[i].Entries.results

            for (var j = 0; j < results.length; j++)
            {
                var option = refiners[i].Name + ":" + results[j].RefinementToken;

                var li = document.createElement("li");
                var a = document.createElement("a");
                a.href = "javascript:;";
                a.innerText = results[j].RefinementName;
                a.className = "";
                a.data = option
                a.addEventListener("click", this.refiner.bind(this), true);

                li.appendChild(a);
                content.appendChild(li);
            }

            refiner.appendChild(spanTitle);
            refiner.appendChild(content);

            refinerList.appendChild(refiner);
        }

        this.appendChild(refinerList);

    },
    loadOrdination: function()
    {
        var ordenationElement = document.createElement("select");

        var arrayOrdenation = this.ordination.split(',');

        for (var i = 0; i < arrayOrdenation.length; i++)
        {
            var ordenationDescending = document.createElement("option");
            ordenationDescending.value = arrayOrdenation[i] + ":descending";
            ordenationDescending.text = arrayOrdenation[i] + " descending";

            var ordenationAscending = document.createElement("option");
            ordenationAscending.value = arrayOrdenation[i] + ":ascending";
            ordenationAscending.text = arrayOrdenation[i] + " ascending";

            ordenationElement.appendChild(ordenationDescending);
            ordenationElement.appendChild(ordenationAscending);
        }

        ordenationElement.addEventListener("change", this.order.bind(this), true);

        this.appendChild(ordenationElement);
    },
    order : function(event)
    {
        var searchText = document.getElementById("searchbox").value;

        var request = new XMLHttpRequest();

        var requestText = "https://" + this.tenant + "-my.sharepoint.com/_api/search/query?querytext='" + searchText + "'&rowlimit=" + this.pageElement;

        if (this.refinement != '')
            requestText += "&refiners='" + this.refinement + "'";

        if (this.refinementFilters != '')
            requestText += "&refinementfilters='" + this.refinementFilters + "'";

        this.sortingList = event.target.value;

        requestText += "&sortlist='" + event.target.value + "'";

        request.open("GET", requestText, true);
        request.setRequestHeader("Authorization", "Bearer " + token);
        request.setRequestHeader("Access-Control-Allow-Origin", "*");
        request.setRequestHeader("accept", "application/json;odata=verbose");
        request.setRequestHeader("content-type", "application/json;odata=verbose");

        request.parent = this

        request.onreadystatechange = function () {
            if (request.readyState == 4 && request.status == 200) {
                var jsonResponseSearch = JSON.parse(request.responseText);

                this.parent.totalElement = jsonResponseSearch.d.query.PrimaryQueryResult.RelevantResults.TotalRows;

                if (this.parent.refinement != '')
                    this.parent.loadRefiner(jsonResponseSearch.d.query.PrimaryQueryResult.RefinementResults.Refiners.results);

                this.parent.showResults(jsonResponseSearch.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results,1);
            }
        }

        request.send();
    },
    refiner: function(event)
    {
        var searchText = document.getElementById("searchbox").value;

        var request = new XMLHttpRequest();

        var requestText = "https://" + this.tenant + "-my.sharepoint.com/_api/search/query?querytext='" + searchText + "'&rowlimit=" + this.pageElement;

        this.sortingList = '';

        if (this.refinement != '')
            requestText += "&refiners='" + this.refinement + "'";

        this.refinementFilters = event.target.data;

        requestText += "&refinementfilters='" + event.target.data + "'";

        request.open("GET", requestText, true);
        request.setRequestHeader("Authorization", "Bearer " + token);
        request.setRequestHeader("Access-Control-Allow-Origin", "*");
        request.setRequestHeader("accept", "application/json;odata=verbose");
        request.setRequestHeader("content-type", "application/json;odata=verbose");

        request.parent = this

        request.onreadystatechange = function () {
            if (request.readyState == 4 && request.status == 200) {
                var jsonResponseSearch = JSON.parse(request.responseText);

                this.parent.totalElement = jsonResponseSearch.d.query.PrimaryQueryResult.RelevantResults.TotalRows;

                if (this.parent.refinement != '')
                    this.parent.loadRefiner(jsonResponseSearch.d.query.PrimaryQueryResult.RefinementResults.Refiners.results);

                this.parent.showResults(jsonResponseSearch.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results,1);
            }
        }

        request.send();
    },
    showResults: function (results,page) 
    {
        var boxResult = document.getElementById("total-results-box");
        var boxToRemoved = document.getElementById("results-box");
        var boxPaging = document.getElementById("page-box");

        if (boxResult != null)
            this.removeChild(boxResult);

        if (boxToRemoved != null)
            this.removeChild(boxToRemoved);

        if (boxPaging != null)
            this.removeChild(boxPaging);

        this.loadTotalElement(page);

        if(this.resultElement != '')
        {
            this.elementShowing(results)
        }
        else
        {
            this.defaultShowing(results)
        }

        this.loadPaging();
    },
    defaultShowing: function(results)
    {
        var resultBox = document.createElement("div");
        resultBox.id = "results-box";
        resultBox.className = "results-box";

        var resultList = document.createElement("ul");

        for (var i = 0; i < results.length; i++)
        {
            var li = document.createElement("li");
            var a = document.createElement("a");
            a.href = results[i].Cells.results[6].Value;
            a.innerText = results[i].Cells.results[3].Value;

            li.appendChild(a);
            resultList.appendChild(li);
        }

        resultBox.appendChild(resultList);
        this.appendChild(resultBox);
    },
    elementShowing: function(results)
    {

    },
    loadTotalElement: function (page)
    {

        var totalBox = document.createElement("div");
        totalBox.id = "total-results-box";
        totalBox.className = "total-results-box";

        var element = parseInt(this.pageElement);

        var initElement = ((page - 1) * element) + 1;
        var endElement = page * element;
        var total = parseInt(this.totalElement);

        if (endElement > total)
            endElement = total;

        var span = document.createElement("span");
        span.innerText = initElement + "-" + endElement + " of " + total + " results";

        totalBox.appendChild(span);

        this.appendChild(totalBox);
    },
    loadPaging: function ()
    {

        var pagingBox = document.createElement("div");
        pagingBox.id = "page-box";
        pagingBox.className = "page-box";
        
        var total = parseInt(this.totalElement);
        var element = parseInt(this.pageElement);

        var pages = (total / element) | 0;

        if (this.totalElement % this.pageElement != 0)
            pages = pages + 1;

        var pagingList = document.createElement("ul");

        var firstPage = document.createElement("li");
        var afirst = document.createElement("a");
        afirst.href = "javascript:;";
        afirst.page = 1;
        afirst.innerText = "First Page";
        afirst.addEventListener("click", this.searchPage.bind(this), true);
        firstPage.appendChild(afirst);
        pagingList.appendChild(firstPage);

        for (var i = 1; i <= pages; i++)
        {
            var page = document.createElement("li");
            var apage = document.createElement("a");
            apage.href = "javascript:;";
            apage.page = i;
            apage.innerText = i;
            apage.addEventListener("click", this.searchPage.bind(this), true);
            page.appendChild(apage);

            pagingList.appendChild(page);
        }


        var lastPage = document.createElement("li");
        var alast = document.createElement("a");
        alast.href = "javascript:;";
        alast.page = pages;
        alast.innerText = "Last Page";
        alast.addEventListener("click", this.searchPage.bind(this), true);
        lastPage.appendChild(alast);
        pagingList.appendChild(lastPage);

        pagingBox.appendChild(pagingList);

        this.appendChild(pagingBox);
    },
    searchPage: function (event)
    {
        var searchText = document.getElementById("searchbox").value;

        var startRow = (event.target.page - 1) * this.pageElement;

        var request = new XMLHttpRequest();

        var requestText = "https://" + this.tenant + "-my.sharepoint.com/_api/search/query?querytext='" + searchText + "'&startrow=" + startRow + "&rowlimit=" + this.pageElement;

        if (this.refinement != '')
            requestText += "&refiners='" + this.refinement + "'";

        if (this.refinementFilters != '')
            requestText += "&refinementfilters='" + this.refinementFilters + "'";

        console.log(this.sortingList);

        if (this.sortingList != '')
            requestText += "&sortlist='" + this.sortingList + "'";

        console.log(requestText);

        request.open("GET", requestText, true);
        request.setRequestHeader("Authorization", "Bearer " + token);
        request.setRequestHeader("Access-Control-Allow-Origin", "*");
        request.setRequestHeader("accept", "application/json;odata=verbose");
        request.setRequestHeader("content-type", "application/json;odata=verbose");

        request.parent = this

        request.onreadystatechange = function () {
            if (request.readyState == 4 && request.status == 200) {
                var jsonResponseSearch = JSON.parse(request.responseText);
                
                if (this.parent.refinement != '')
                    this.parent.loadRefiner(jsonResponseSearch.d.query.PrimaryQueryResult.RefinementResults.Refiners.results);

                this.parent.showResults(jsonResponseSearch.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results, event.target.page);
            }
        }

        request.send();
    }
});
