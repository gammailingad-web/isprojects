Branch Queuing Launchers - GitHub URL update

Updated HTML files: 29

All branch launcher files now read:
https://raw.githubusercontent.com/gammailingad-web/ISProjects-URL-Manager/main/isportal-url.txt

Behavior:
- Fetches the public GitHub Raw text file.
- Uses ONLY the first line as the current ISProjects Cloudflare URL.
- Ignores the second line (COC Portal URL).
- Validates that line 1 is an HTTPS trycloudflare.com URL.
- Preserves each launcher's original queuing destination and branch code.

Installation:
1. Back up your existing branch launcher HTML files.
2. Extract the HTML files from this ZIP.
3. Replace the corresponding existing branch launcher files.
4. Keep the GitHub repository public.

No GitHub token is needed in these HTML files.
