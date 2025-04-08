export const ensureNamespaces = (docElement, nsMap) => {
  for (const [prefix, uri] of Object.entries(nsMap)) {
    if (!docElement.hasAttribute(`xmlns:${prefix}`)) {
      docElement.setAttribute(`xmlns:${prefix}`, uri);
    }
  }
}
