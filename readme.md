# word-doc-var

To support HTML content as a value in content control.

## Usage

1. Make sure to register the `HtmlControlReplacer` in the engine. 

```cs
var engine = new DefaultOpenXmlTemplateEngine();

engine.RegisterReplacer(new HtmlControlReplacer());
```

2. the template variable's tag should be prefixed with `html_`

3. the template's content control type should be of rich text.