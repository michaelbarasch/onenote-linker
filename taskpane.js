Office.onReady((info) => {
    if (info.host === Office.HostType.OneNote) {
        document.getElementById("btnGenerate").onclick = generateLinkList;
    }
});

async function generateLinkList() {
    const searchTerm = document.getElementById("searchInput").value.trim().toLowerCase();
    const status = document.getElementById("status");

    if (!searchTerm) {
        status.innerText = "Error: Please enter a title to search.";
        return;
    }

    status.innerText = "Searching...";

    try {
        await OneNote.run(async (context) => {
            // 1. Get all pages in the current notebook
            const notebook = context.application.getActiveNotebook();
            const pages = notebook.getPages();
            
            // 2. Load necessary properties (clientUrl is best for Desktop app)
            pages.load("title, clientUrl");
            await context.sync();

            // 3. Filter for matches
            const matches = pages.items.filter(page => 
                page.title.toLowerCase().includes(searchTerm)
            );

            if (matches.length === 0) {
                status.innerText = "No matching pages found.";
                return;
            }

            // 4. Create the HTML list string
            let htmlContent = "<div><strong>Search Results:</strong><ul>";
            matches.forEach(page => {
                // Formatting as a standard HTML anchor tag
                htmlContent += `<li><a href="${page.clientUrl}">${page.title}</a></li>`;
            });
            htmlContent += "</ul></div>";

            // 5. Insert onto the current active page
            const activePage = context.application.getActivePage();
            activePage.addOutline(100, 100, htmlContent);

            await context.sync();
            status.innerText = `Success! Added ${matches.length} links.`;
        });
    } catch (error) {
        status.innerText = "Error: " + error.message;
        console.error(error);
    }
}