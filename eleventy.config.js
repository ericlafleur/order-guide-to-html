const path = require("path");
const fs = require("fs");

module.exports = function (eleventyConfig) {
  // Copy all HTML files from workbook_html/ to _site/workbook_html/ as-is.
  eleventyConfig.addPassthroughCopy("workbook_html");

  // Build a collection of workbook pages so the sitemap knows about them.
  eleventyConfig.addCollection("workbookPages", function (_collectionApi) {
    const pages = [];

    function walkDir(currentPath) {
      if (!fs.existsSync(currentPath)) return;
      for (const entry of fs.readdirSync(currentPath, {
        withFileTypes: true,
      })) {
        const fullPath = path.join(currentPath, entry.name);
        if (entry.isDirectory()) {
          walkDir(fullPath);
        } else if (entry.name.endsWith(".html")) {
          const mtime = fs.statSync(fullPath).mtime;
          pages.push({
            url: "/" + fullPath.split(path.sep).join("/"),
            date: mtime.toISOString().split("T")[0],
          });
        }
      }
    }

    walkDir("workbook_html");
    return pages;
  });

  // Prevent Eleventy from processing the README.
  eleventyConfig.ignores.add("README.md");

  return {
    // Only treat .njk and .md files as templates; HTML files in workbook_html/
    // are passthrough-copied as-is via addPassthroughCopy above.
    templateFormats: ["njk", "md"],
    dir: {
      input: ".",
      output: "_site",
    },
  };
};
