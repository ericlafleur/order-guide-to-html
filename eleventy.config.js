const path = require("path");
const fs = require("fs");

module.exports = function (eleventyConfig) {
  // Copy all HTML files from workbooks_html/ to _site/workbooks_html/ as-is.
  eleventyConfig.addPassthroughCopy("workbooks_html");

  // Build a metadata lookup from all manifest_*.json files in workbooks_html/.
  const metadataByFilename = {};
  const htmlDir = "workbooks_html";
  if (fs.existsSync(htmlDir)) {
    for (const entry of fs.readdirSync(htmlDir)) {
      if (entry.startsWith("manifest_") && entry.endsWith(".json")) {
        try {
          const manifest = JSON.parse(
            fs.readFileSync(path.join(htmlDir, entry), "utf-8")
          );
          for (const fileEntry of manifest.files || []) {
            if (fileEntry.path) {
              const filename = path.basename(fileEntry.path);
              metadataByFilename[filename] = {
                vehicle_name: manifest.vehicle_name,
                ...fileEntry,
              };
            }
          }
        } catch (_e) {
          // skip unreadable or malformed manifest files
        }
      }
    }
  }

  // Build a collection of workbook pages for a given language subdirectory.
  function buildWorkbookPagesCollection(langDir) {
    return function (_collectionApi) {
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
              meta: metadataByFilename[entry.name] || null,
            });
          }
        }
      }

      walkDir(langDir);
      return pages;
    };
  }

  eleventyConfig.addCollection("workbookPagesEn", buildWorkbookPagesCollection("workbooks_html/en"));
  eleventyConfig.addCollection("workbookPagesFr", buildWorkbookPagesCollection("workbooks_html/fr"));

  // Prevent Eleventy from processing the README.
  eleventyConfig.ignores.add("README.md");

  return {
    // Only treat .njk and .md files as templates; HTML files in workbooks_html/
    // are passthrough-copied as-is via addPassthroughCopy above.
    templateFormats: ["njk", "md"],
    dir: {
      input: ".",
      output: "_site",
    },
  };
};
