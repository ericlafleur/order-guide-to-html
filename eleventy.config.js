const path = require("path");
const fs = require("fs");

module.exports = function (eleventyConfig) {
  // Copy all HTML files from workbooks_html/ to _site/workbooks_html/ as-is.
  eleventyConfig.addPassthroughCopy("workbooks_html");

  // Build a metadata lookup from all manifest_*.json files under workbooks_html/.
  const metadataByFilename = {};
  const metadataByRelativePath = {};
  const htmlDir = "workbooks_html";

  function walkDir(currentPath, onFile) {
    if (!fs.existsSync(currentPath)) return;
    for (const entry of fs.readdirSync(currentPath, { withFileTypes: true })) {
      const fullPath = path.join(currentPath, entry.name);
      if (entry.isDirectory()) {
        walkDir(fullPath, onFile);
      } else {
        onFile(fullPath, entry.name);
      }
    }
  }

  if (fs.existsSync(htmlDir)) {
    walkDir(htmlDir, (fullPath, entryName) => {
      if (entryName.startsWith("manifest_") && entryName.endsWith(".json")) {
        try {
          const manifest = JSON.parse(
            fs.readFileSync(fullPath, "utf-8")
          );
          const manifestDir = path.dirname(fullPath).split(path.sep).join("/");
          for (const fileEntry of manifest.files || []) {
            if (fileEntry.path) {
              const filename = path.basename(fileEntry.path);
              const relativeFilePath = manifestDir + "/" + filename;
              const metadata = {
                vehicle_name: manifest.vehicle_name,
                ...fileEntry,
              };
              metadataByRelativePath[relativeFilePath] = metadata;
              // Keep filename fallback for legacy flat outputs (first writer wins).
              if (!metadataByFilename[filename]) {
                metadataByFilename[filename] = metadata;
              }
            }
          }
        } catch (_e) {
          // skip unreadable or malformed manifest files
        }
      }
    });
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
            const relativePath = fullPath.split(path.sep).join("/");
            pages.push({
              url: "/" + relativePath,
              date: mtime.toISOString().replace(/\.\d{3}Z$/, "Z"),
              meta:
                metadataByRelativePath[relativePath] ||
                metadataByFilename[entry.name] ||
                null,
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
