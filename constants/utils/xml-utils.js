export const ensureNamespaces = (element, namespaces) => {
  const attrs = new Set();

  for (let i = 0; i < element.attributes.length; i++) {
    const attr = element.attributes[i];
    if (attr.name.startsWith("xmlns:")) {
      attrs.add(attr.name.replace("xmlns:", ""));
    }
  }

  for (const [prefix, uri] of Object.entries(namespaces)) {
    if (!attrs.has(prefix)) {
      element.setAttribute(`xmlns:${prefix}`, uri);
    }
  }
};
