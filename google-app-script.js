/**
 * GoogleDocumentToHTML
 * Author: G.H. Bos
 * 
 * https://github.com/Gijsbos/GoogleDocumentToHTML
 * 
 * Usage:
 *  - Set output to 'print' or 'email' in GoogleDocumentToHTML(output = "<output type>")
 * 
 * Adaptation of https://github.com/oazabir/GoogleDoc2Html:
 *  - Improved html formatting (indentation etc)
 *  - Fixed list element nesting
 */

/**
 * GoogleDocumentToHTML
 */
function GoogleDocumentToHTML(output = "print") // print/email
{
  main(output);
}

/**
 * Utils
 */
open = (tag) => { return `<${tag}>`; };
close = (tag) => { return `</${tag}>`; };
indent = (indents) => { return `\t`.repeat(indents); };
html = (tag, content, indents, newline = true) => { return `${newline ? '\n' : ''}${indent(indents)}${open(tag)}\n${content}\n${indent(indents)}${close(tag)}`};

/**
 * GoogleDocumentToHTML
 */
function main(output)
{
  // Create GoogleElementTree
  var googleDocument = new GoogleElementTree();

  // Get body
  var body = DocumentApp.getActiveDocument().getBody();

  // Walk through children
  for (var i = 0; i < body.getNumChildren(); i++)
    googleDocument.add(body.getChild(i));

  // Set output: print or email content to self (required when html is too. large)
  if(output == "email")
    googleDocument.emailHTML();
  else
    googleDocument.printHTML();
}

/**
 * GoogleElementTree
 */
class GoogleElementTree
{
  /**
   * constructor
   */
  constructor()
  {
    // Prepare tree
    this.items = [];

    // Used during export
    this.images = [];
  }

  /**
   * add
   */
  add(item, nestingLevel = 0, items = undefined)
  {
    // Init list
    items = items === undefined ? this.items : items;

    // Check item type
    if(item.getListId)
    {
      let lastItem = items[items.length - 1];

      if(Array.isArray(lastItem))
      {
        if(nestingLevel == parseInt(item.getNestingLevel()))
        {
          lastItem.push(item);
        }
        else
        {
          this.add(item, nestingLevel + 1, lastItem);
        }
      }
      else
      {
        items.push([item]);
      }
    }
    else
    {
      items.push(item);
    }
  }

  /**
   * getParagraphTag
   */
  getParagraphTag(item)
  {
    switch(item.getHeading())
    {
      case DocumentApp.ParagraphHeading.HEADING6: return "h6";
      case DocumentApp.ParagraphHeading.HEADING5: return "h5";
      case DocumentApp.ParagraphHeading.HEADING4: return "h4";
      case DocumentApp.ParagraphHeading.HEADING3: return "h3";
      case DocumentApp.ParagraphHeading.HEADING2: return "h2";
      case DocumentApp.ParagraphHeading.HEADING1: return "h1";
      default: return "p";
    }
  }

  /**
   * getOpenClose
   */
  getTag(item)
  {
    switch(item.getType())
    {
      case DocumentApp.ElementType.PARAGRAPH:
        return this.getParagraphTag(item);
      case DocumentApp.ElementType.LIST_ITEM:
        return "li";
      default:
        return false;
    }
  }
  
  /**
   * parseTextChild
   */
  parseTextChild(item)
  {
    var output = [];
    var text = item.getText();
    var indices = item.getTextAttributeIndices();

    // Replace symbols
    text = text.replace(/‘|’/g,"'");

    // Proceed
    for (var i= 0; i < indices.length; i++)
    {
      var partAtts = item.getAttributes(indices[i]);
      var startPos = indices[i];
      var endPos = i+1 < indices.length ? indices[i+1]: text.length;
      var partText = text.substring(startPos, endPos);

      // Add opening tag
      if (partAtts.ITALIC)
        output.push('<i>');

      if (partAtts.BOLD)
        output.push('<b>');

      if (partAtts.UNDERLINE)
        output.push('<u>');

      if (partAtts.LINK_URL)
        output.push(`<a href="${partText.indexOf("http") === 0 ? partText : 'http://' + partText}">`);

      // Add text
      output.push(partText);

      // Add closing tag
      if (partAtts.LINK_URL)
        output.push(`</a>`);

      if (partAtts.UNDERLINE)
        output.push('</u>');

      if (partAtts.BOLD)
        output.push('</b>');

      if (partAtts.ITALIC)
        output.push('</i>');
    }

    return output.join("");
  }

  /**
   * parseText
   */
  parseText(item)
  {
    if(item.getNumChildren())
      return this.parseTextChild(item.getChild(0));

    return item.getText();
  }

  /**
   * processImage
   */
  processImage(item)
  {
    var blob = item.getBlob();
    var contentType = blob.getContentType();

    var extension = "";
    if (/\/png$/.test(contentType))
      extension = ".png";
    else if (/\/gif$/.test(contentType))
      extension = ".gif";
    else if (/\/jpe?g$/.test(contentType))
      extension = ".jpg";
    else
      throw "Unsupported image type: "+contentType;

    // Name
    var name = "Image_" + this.images.length + extension;

    // Add image
    this.images.push({
      "blob": blob,
      "type": contentType,
      "name": name
    });

    // Return html
    return `<img src="cid:${name}" />`;
  }

  /**
   * exportItem
   */
  exportItem(item, parent, indents)
  {
    if(item.getType() == DocumentApp.ElementType.INLINE_IMAGE)
      return this.processImage(item);
    else
    {
      let tag = this.getTag(item);
      let text = this.parseText(item);

      // Line breaks are empty <p> elements, replace with </br>
      if(tag === 'p' && text.trim().length === 0)
        return "\n</br>";

      // Parent is array
      if(Array.isArray(parent))
      {
        // Prevent inserting a newline for first item in parent
        if(parent.indexOf(item) == 0)
          return html(tag, `${indent(indents+1)}${text}`, indents, false);
      }

      return html(tag, `${indent(indents+1)}${text}`, indents);
    }
  }

  /**
   * getListTag
   */
  getListTag(item)
  {
    switch(item.getGlyphType())
    {
      case DocumentApp.GlyphType.BULLET:
      case DocumentApp.GlyphType.HOLLOW_BULLET:
      case DocumentApp.GlyphType.SQUARE_BULLET:
        return "ul";
      default:
        return "ol";
    }
  }

  /**
   * exportToHTML
   */
  exportToHTML(items = undefined, indents = 0)
  {
    items = items === undefined ? this.items : items;

    // Set content
    let content = [];

    // Iterate over items
    for(let item of items)
    {
      // List
      if(Array.isArray(item))
      {
        content.push(
          html(
            this.getListTag(item[0]),
            this.exportToHTML(item, indents + 1),
            indents
          )
        )
      }

      // Not a list
      else
        content.push(this.exportItem(item, items, indents));
    }

    // Return string
    return content.join("");
  }

  /**
   * printHTML
   */
  printHTML()
  {
    Logger.log(this.exportToHTML());
  }

  /**
   * emailHTML
   */
  emailHTML()
  {
    var html = this.exportToHTML();
    var images = this.images;    
    var attachments = [];

    // Process images
    for (var j=0; j<images.length; j++)
    {
      attachments.push({
        "fileName": images[j].name,
        "mimeType": images[j].type,
        "content": images[j].blob.getBytes()
      });
    }

    var inlineImages = {};
    for (var j=0; j<images.length; j++)
      inlineImages[[images[j].name]] = images[j].blob;

    // Name
    var name = DocumentApp.getActiveDocument().getName()+".html";

    // Add attachment
    attachments.push({"fileName":name, "mimeType": "text/html", "content": html});

    // Send mail
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: name,
      htmlBody: html,
      inlineImages: inlineImages,
      attachments: attachments
    });
  }
}
