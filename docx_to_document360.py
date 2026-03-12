from __future__ import annotations

import argparse
import base64
from datetime import datetime
import json
import html
import os
import re
import sys
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import mammoth
import requests
from bs4 import BeautifulSoup, Comment


DEFAULT_BASE_URL = "https://apihub.document360.io"
DEFAULT_CATEGORY_NAME = "Imports"
DEFAULT_PORTAL_URL = ""
STYLE_MAP = """
p[style-name='Title'] => h1:fresh
p[style-name='Heading 1'] => h1:fresh
p[style-name='Heading 2'] => h2:fresh
p[style-name='Heading 3'] => h3:fresh
p[style-name='Heading 4'] => h4:fresh
p[style-name='Heading 5'] => h5:fresh
p[style-name='Heading 6'] => h6:fresh
p[style-name='Code'] => pre:fresh
r[style-name='Strong'] => strong
""".strip()


def load_dotenv(dotenv_path: str = ".env") -> None:
    path = Path(dotenv_path)
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'").strip('"')
        if key and key not in os.environ:
            os.environ[key] = value


class Document360Client:
    def __init__(self, token: str, base_url: str) -> None:
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update(
            {
                "api_token": token,
                "Content-Type": "application/json",
                "Accept": "application/json",
            }
        )

    def _request(self, method: str, path: str, **kwargs: Any) -> dict[str, Any]:
        response = self.session.request(method, f"{self.base_url}{path}", timeout=60, **kwargs)
        print(f"{method.upper()} {path} -> {response.status_code}")
        if not response.ok:
            raise RuntimeError(
                f"Document360 API error {response.status_code} for {path}: {response.text[:1000]}"
            )
        if not response.text.strip():
            return {}
        return response.json()

    def list_project_versions(self) -> list[dict[str, Any]]:
        payload = self._request("GET", "/v2/ProjectVersions")
        return payload.get("data") or []

    def list_team_accounts(self) -> list[dict[str, Any]]:
        payload = self._request("GET", "/v2/Teams", params={"skip": 0, "take": 100})
        return payload.get("result") or []

    def list_languages(self, project_version_id: str) -> list[dict[str, Any]]:
        payload = self._request("GET", f"/v2/Language/{project_version_id}")
        return payload.get("data") or []

    def list_categories(self, project_version_id: str, lang_code: str) -> list[dict[str, Any]]:
        payload = self._request(
            "GET",
            f"/v2/ProjectVersions/{project_version_id}/categories",
            params={
                "excludeArticles": True,
                "includeCategoryDescription": False,
                "langCode": lang_code,
            },
        )
        return payload.get("data") or []

    def create_category(
        self,
        *,
        name: str,
        project_version_id: str,
        user_id: str,
        parent_category_id: str | None = None,
    ) -> dict[str, Any]:
        payload = self._request(
            "POST",
            "/v2/Categories",
            data=json.dumps(
                {
                    "name": name,
                    "project_version_id": project_version_id,
                    "order": 0,
                    "parent_category_id": parent_category_id,
                    "content": None,
                    "category_type": 0,
                    "user_id": user_id,
                    "content_type": None,
                    "slug": None,
                }
            ),
        )
        return payload.get("data") or payload

    def create_article(
        self,
        *,
        title: str,
        html: str,
        category_id: str,
        project_version_id: str,
        user_id: str,
        slug: str | None = None,
    ) -> dict[str, Any]:
        payload = self._request(
            "POST",
            "/v2/Articles",
            data=json.dumps(
                {
                    "title": title,
                    "content": html,
                    "category_id": category_id,
                    "project_version_id": project_version_id,
                    "order": 0,
                    "user_id": user_id,
                    "content_type": 1,
                    "article_type": None,
                    "slug": slug,
                }
            ),
        )
        return payload.get("data") or payload

    def publish_article(
        self,
        *,
        article_id: str,
        lang_code: str,
        user_id: str,
        version_number: int = 1,
        publish_message: str = "Initial publish from DOCX import",
    ) -> dict[str, Any]:
        payload = self._request(
            "POST",
            f"/v2/Articles/{article_id}/{lang_code}/publish",
            data=json.dumps(
                {
                    "user_id": user_id,
                    "version_number": version_number,
                    "publish_message": publish_message,
                }
            ),
        )
        return payload


def convert_image(image: mammoth.images.Image) -> dict[str, str]:
    with image.open() as image_bytes:
        encoded = base64.b64encode(image_bytes.read()).decode("ascii")
    content_type = image.content_type or "application/octet-stream"
    return {"src": f"data:{content_type};base64,{encoded}", "alt": image.alt_text or ""}


def convert_docx_to_html(docx_path: Path) -> tuple[str, str]:
    with docx_path.open("rb") as docx_file:
        result = mammoth.convert_to_html(
            docx_file,
            style_map=STYLE_MAP,
            convert_image=mammoth.images.img_element(convert_image),
        )
    html = clean_html(result.value)
    title = extract_title(html, docx_path)
    html = strip_duplicate_title_heading(html, title)
    return title, html


def clean_html(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    for comment in soup.find_all(string=lambda value: isinstance(value, Comment)):
        comment.extract()

    for tag in soup.find_all(["script", "style"]):
        tag.decompose()

    for tag in soup.find_all(True):
        if tag.attrs:
            allowed = {}
            if tag.name == "a" and tag.get("href"):
                allowed["href"] = tag["href"]
            if tag.name == "img":
                if tag.get("src"):
                    allowed["src"] = tag["src"]
                if tag.get("alt"):
                    allowed["alt"] = tag["alt"]
            tag.attrs = allowed

    for tag in soup.find_all(["p", "li", "td", "th"]):
        text = tag.get_text(" ", strip=True)
        if tag.name == "p" and not text and not tag.find("img"):
            tag.decompose()

    for table in soup.find_all("table"):
        first_row = table.find("tr")
        if first_row and first_row.find("td") and not first_row.find("th"):
            for cell in first_row.find_all("td"):
                cell.name = "th"

    for cell in soup.find_all(["td", "th"]):
        paragraphs = cell.find_all("p", recursive=False)
        if len(paragraphs) == 1 and not paragraphs[0].attrs:
            paragraphs[0].unwrap()

    for anchor in soup.find_all("a"):
        href = (anchor.get("href") or "").strip()
        if not href:
            anchor.unwrap()
            continue
        parsed = urlparse(href)
        if parsed.scheme not in {"http", "https", "mailto"}:
            anchor.unwrap()

    for paragraph in soup.find_all("p"):
        if paragraph.find("br") and is_code_like(paragraph):
            pre = soup.new_tag("pre")
            pre.string = paragraph.get_text("\n", strip=False).strip("\n")
            paragraph.replace_with(pre)

    for paragraph in soup.find_all("p"):
        linkified = linkify_plain_urls(paragraph.get_text("\n", strip=False))
        if linkified is not None:
            replacement = BeautifulSoup(linkified, "html.parser")
            paragraph.clear()
            for child in list(replacement.contents):
                paragraph.append(child)

    for pre in soup.find_all("pre"):
        text = pre.get_text("\n", strip=False)
        pre.clear()
        pre.string = text.strip("\n")

    for image in soup.find_all("img"):
        src = image.get("src", "")
        if src.startswith("data:") and len(src) > 200_000:
            image.replace_with(image.get("alt") or "Embedded image omitted from import.")

    for paragraph in soup.find_all("p"):
        text = normalize_space(paragraph.get_text(" ", strip=True))
        if re.fullmatch(r"Figure\s*", text):
            paragraph.decompose()

    body = "".join(str(node) for node in soup.contents).strip()
    body = re.sub(r"\n{3,}", "\n\n", body)
    return body


def extract_title(html: str, docx_path: Path) -> str:
    soup = BeautifulSoup(html, "html.parser")
    heading = soup.find(["h1", "h2"])
    if heading:
        title = heading.get_text(" ", strip=True)
        if title:
            return title

    paragraph = soup.find(["p", "li"])
    if paragraph:
        title = paragraph.get_text(" ", strip=True)
        if title:
            return truncate_title(title)

    return docx_path.stem.replace("_", " ").strip() or "Imported Document"


def strip_duplicate_title_heading(html: str, title: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    first_heading = soup.find(["h1", "h2"])
    if first_heading and first_heading.get_text(" ", strip=True) == title:
        first_heading.decompose()
    return "".join(str(node) for node in soup.contents).strip()


def truncate_title(value: str, limit: int = 120) -> str:
    compact = re.sub(r"\s+", " ", value).strip()
    if len(compact) <= limit:
        return compact
    return compact[: limit - 1].rstrip() + "…"


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return slug[:80] or "imported-document"


def unique_slug(base_slug: str) -> str:
    suffix = datetime.now().strftime("%Y%m%d%H%M%S")
    candidate = f"{base_slug}-{suffix}"
    return candidate[:120]


def flatten_categories(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for item in items:
        output.append(item)
        children = item.get("child_categories") or item.get("categories") or []
        output.extend(flatten_categories(children))
    return output


def is_code_like(paragraph: Any) -> bool:
    text = paragraph.get_text("\n", strip=False)
    lines = [line.rstrip() for line in text.splitlines() if line.strip()]
    if len(lines) < 2:
        return False
    score = 0
    for line in lines:
        if re.search(r"[{}();=\[\]]", line):
            score += 1
        if line.startswith(("    ", "\t")):
            score += 1
        if re.match(r"(public|private|protected|class|if|for|while)\b", line.strip()):
            score += 1
    return score >= 2


def linkify_plain_urls(text: str) -> str | None:
    pattern = re.compile(r"(?P<url>https?://[^\s<]+)")
    if not pattern.search(text):
        return None

    def replace(match: re.Match[str]) -> str:
        url = match.group("url")
        escaped = html.escape(url, quote=True)
        return f'<a href="{escaped}">{escaped}</a>'

    return pattern.sub(replace, html.escape(text))


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def build_article_url(
    *,
    portal_url: str,
    slug: str,
    article: dict[str, Any],
) -> str | None:
    base = portal_url.rstrip("/")
    if not base:
        return None

    url = article.get("url") or article.get("article_url")
    if isinstance(url, str) and url.strip():
        if url.startswith(("http://", "https://")):
            return url
        return f"{base}/{url.lstrip('/')}"

    return f"{base}/docs/{slug}"


def choose_project_version(client: Document360Client, project_version_id: str | None) -> dict[str, Any]:
    versions = client.list_project_versions()
    if not versions:
        raise RuntimeError("No Document360 project versions were returned by the API.")

    if project_version_id:
        for version in versions:
            if version.get("id") == project_version_id:
                return version
        raise RuntimeError(f"Configured project version was not found: {project_version_id}")

    main_version = next((item for item in versions if item.get("is_main_version")), None)
    return main_version or versions[0]


def choose_user(client: Document360Client, user_id: str | None) -> dict[str, Any]:
    users = client.list_team_accounts()
    if not users:
        raise RuntimeError("No team accounts were returned by the API.")

    if user_id:
        for user in users:
            if user.get("user_id") == user_id:
                return user
        raise RuntimeError(f"Configured user_id was not found: {user_id}")

    return users[0]


def choose_lang_code(
    client: Document360Client,
    project_version_id: str,
    lang_code: str | None,
) -> str:
    if lang_code:
        return lang_code

    languages = client.list_languages(project_version_id)
    default_language = next((item for item in languages if item.get("is_set_as_default")), None)
    if default_language and default_language.get("language_code"):
        return default_language["language_code"]
    return "en"


def choose_or_create_category(
    client: Document360Client,
    *,
    project_version_id: str,
    lang_code: str,
    user_id: str,
    category_id: str | None,
    category_name: str,
) -> dict[str, Any]:
    categories = flatten_categories(client.list_categories(project_version_id, lang_code))

    if category_id:
        for category in categories:
            if category.get("id") == category_id:
                return category
        raise RuntimeError(f"Configured category_id was not found: {category_id}")

    for category in categories:
        if (category.get("name") or "").strip().lower() == category_name.strip().lower():
            return category

    return client.create_category(
        name=category_name,
        project_version_id=project_version_id,
        user_id=user_id,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert a DOCX file to clean HTML and upload it to Document360."
    )
    parser.add_argument(
        "--docx",
        default=r"c:\Users\pc\Downloads\docx_migration_test_file.docx",
        help="Path to the source DOCX file.",
    )
    parser.add_argument(
        "--output-html",
        default="output.html",
        help="Where to save the converted HTML.",
    )
    parser.add_argument(
        "--title",
        default=None,
        help="Optional override for the article title.",
    )
    parser.add_argument(
        "--no-upload",
        action="store_true",
        help="Only generate HTML locally and skip the Document360 upload.",
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help="Publish the article after creating it.",
    )
    parser.add_argument(
        "--list-targets",
        action="store_true",
        help="List available Document360 project versions, users, languages, and categories.",
    )
    return parser.parse_args()


def print_targets(client: Document360Client, project_version_id: str | None) -> None:
    project_version = choose_project_version(client, project_version_id)
    user = choose_user(client, os.getenv("DOCUMENT360_USER_ID"))
    lang_code = choose_lang_code(client, project_version["id"], os.getenv("DOCUMENT360_LANG_CODE"))
    categories = flatten_categories(client.list_categories(project_version["id"], lang_code))

    print("Project version:")
    print(f"  id={project_version.get('id')} name={project_version.get('name')}")
    print("User:")
    print(f"  user_id={user.get('user_id')} name={user.get('name') or user.get('email')}")
    print("Language:")
    print(f"  code={lang_code}")
    print("Categories:")
    for category in categories:
        print(f"  id={category.get('id')} name={category.get('name')}")


def main() -> int:
    load_dotenv()
    args = parse_args()
    docx_path = Path(args.docx)
    if not docx_path.exists():
        print(f"DOCX file not found: {docx_path}", file=sys.stderr)
        return 1

    title, html = convert_docx_to_html(docx_path)
    if args.title:
        title = args.title

    output_path = Path(args.output_html)
    output_path.write_text(html, encoding="utf-8")
    print(f"Saved HTML to {output_path.resolve()}")
    print(f"Detected title: {title}")

    if args.no_upload:
        return 0

    token = os.getenv("DOCUMENT360_API_TOKEN")
    if not token:
        print("Skipping upload: DOCUMENT360_API_TOKEN is not set.", file=sys.stderr)
        return 2

    base_url = os.getenv("DOCUMENT360_BASE_URL", DEFAULT_BASE_URL)
    category_name = os.getenv("DOCUMENT360_CATEGORY_NAME", DEFAULT_CATEGORY_NAME)
    client = Document360Client(token=token, base_url=base_url)

    if args.list_targets:
        print_targets(client, os.getenv("DOCUMENT360_PROJECT_VERSION_ID"))
        return 0

    project_version = choose_project_version(client, os.getenv("DOCUMENT360_PROJECT_VERSION_ID"))
    user = choose_user(client, os.getenv("DOCUMENT360_USER_ID"))
    lang_code = choose_lang_code(client, project_version["id"], os.getenv("DOCUMENT360_LANG_CODE"))
    category = choose_or_create_category(
        client,
        project_version_id=project_version["id"],
        lang_code=lang_code,
        user_id=user["user_id"],
        category_id=os.getenv("DOCUMENT360_CATEGORY_ID"),
        category_name=category_name,
    )

    article_slug = unique_slug(slugify(title))
    article = client.create_article(
        title=title,
        html=html,
        category_id=category["id"],
        project_version_id=project_version["id"],
        user_id=user["user_id"],
        slug=article_slug,
    )
    article_id = article.get("id")
    if not article_id:
        raise RuntimeError(f"Article creation succeeded but no article id was returned: {article}")

    print(f"Created article: {article_id}")
    article_url = build_article_url(
        portal_url=os.getenv("DOCUMENT360_PORTAL_URL", DEFAULT_PORTAL_URL),
        slug=article_slug,
        article=article,
    )
    if article_url:
        print(f"Article URL: {article_url}")

    should_publish = args.publish or os.getenv("DOCUMENT360_PUBLISH", "true").lower() == "true"
    if should_publish:
        version_number = int(article.get("version_number") or 1)
        client.publish_article(
            article_id=article_id,
            lang_code=lang_code,
            user_id=user["user_id"],
            version_number=version_number,
        )
        print(f"Published article {article_id} in language {lang_code}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
