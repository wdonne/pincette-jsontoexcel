# Merge JSON With Excel

With this library JSON can be merged with an Excel template. There are two cases. You can either provide a single object or a list of objects. The template should contain cells with occurrences of bindings, where one cell can have several bindings. A binding is string of the form `{field}`. A field is a dot-separated path.

When a list of objects is provided the template should have one row in which the cells contain bindings. The template row will be repeated for every object in the list. In the case of a single object there may be bindings in any cell.

When no value is found for a binding, it will be replaced with an empty string.

The class `net.pincette.jsontoexcel.Merge` has the following methods:

```
public static void merge(
      final JsonParser parser, final InputStream template, final OutputStream out)
```

```
public static void merge(
      final JsonObject json, final InputStream template, final OutputStream out)
```

```
public static void merge(
      final Stream<JsonValue> stream, final InputStream template, final OutputStream out)
```

The first one will use one of the two others depending on what is in the stream.