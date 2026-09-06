import { Fragment } from "react";
import { parseMarkdown, type Inline } from "../markdown";

/**
 * Renders the small Markdown subset the model uses. Spans become React
 * elements, so model output is never injected as markup.
 */
export function Markdown({ text }: { text: string }) {
  const blocks = parseMarkdown(text);

  if (blocks.length === 0) {
    return null;
  }

  return (
    <>
      {blocks.map((block, index) => {
        switch (block.type) {
          case "heading":
            return (
              <h3 key={index} className="md-heading">
                <Spans spans={block.spans} />
              </h3>
            );
          case "list":
            return (
              <ul key={index} className="md-list">
                {block.items.map((item, itemIndex) => (
                  <li key={itemIndex}>
                    <Spans spans={item} />
                  </li>
                ))}
              </ul>
            );
          case "rule":
            return <hr key={index} className="md-rule" />;
          default:
            return (
              <p key={index}>
                <Spans spans={block.spans} />
              </p>
            );
        }
      })}
    </>
  );
}

function Spans({ spans }: { spans: Inline[] }) {
  return (
    <>
      {spans.map((span, index) => {
        switch (span.type) {
          case "bold":
            return <strong key={index}>{span.text}</strong>;
          case "italic":
            return <em key={index}>{span.text}</em>;
          case "code":
            return (
              <code key={index} className="md-code">
                {span.text}
              </code>
            );
          default:
            return <Fragment key={index}>{span.text}</Fragment>;
        }
      })}
    </>
  );
}
