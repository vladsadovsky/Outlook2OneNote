// Best-effort HTML sanitisation for OneNote OOXML.
// Strips scripts, iframes, event attributes, and external images.
// No imports from outside src/onenote/ (boundary rule — DESIGN.md 6.1).

const DANGEROUS_TAGS = ['script', 'iframe', 'object', 'embed', 'form', 'base', 'meta', 'link']

// Attributes that can carry JavaScript
const DANGEROUS_ATTRS_RE = /^on[a-z]/i

// CSS values that execute code
const DANGEROUS_CSS_RE = /expression\s*\(|javascript\s*:|behavior\s*:/i

function removeElement(el: Element): void {
  el.parentNode?.removeChild(el)
}

function sanitizeElement(el: Element): void {
  // Work on a snapshot so we can remove while iterating
  for (const attr of Array.from(el.attributes)) {
    if (DANGEROUS_ATTRS_RE.test(attr.name)) {
      el.removeAttribute(attr.name)
      continue
    }
    if (attr.name === 'style' && DANGEROUS_CSS_RE.test(attr.value)) {
      el.removeAttribute('style')
      continue
    }
    // Strip javascript: hrefs/srcs
    if ((attr.name === 'href' || attr.name === 'src' || attr.name === 'action') &&
        /^\s*javascript\s*:/i.test(attr.value)) {
      el.removeAttribute(attr.name)
    }
  }

  // Remove external images — OneNote cannot reliably render them and they are a privacy risk
  if (el.tagName.toLowerCase() === 'img') {
    const src = el.getAttribute('src') ?? ''
    if (/^https?:\/\//i.test(src)) {
      // Replace with a note rather than removing the element entirely, to preserve layout
      el.removeAttribute('src')
      el.setAttribute('alt', el.getAttribute('alt') ?? '[external image removed]')
    }
  }
}

export function sanitizeHtml(html: string): string {
  const parser = new DOMParser()
  const doc = parser.parseFromString(html, 'text/html')

  // Remove dangerous tags entirely (reverse order so children are removed before parents)
  for (const tag of DANGEROUS_TAGS) {
    for (const el of Array.from(doc.querySelectorAll(tag))) {
      removeElement(el)
    }
  }

  // Walk all remaining elements and sanitize attributes
  for (const el of Array.from(doc.body.querySelectorAll('*'))) {
    sanitizeElement(el)
  }

  return doc.body.innerHTML
}
