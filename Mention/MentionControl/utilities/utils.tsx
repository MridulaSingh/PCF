export const getRangeAtCursorPosition = () => {
    if (!window.getSelection) {
        return null
    }

    const sel = window.getSelection();
    if (sel != null && (!sel.getRangeAt || !sel.rangeCount)) {
        return null
    }
    else if (sel!=null){

        return sel.getRangeAt(0)
    }
}

export const insertSpanAtCursorPosition = (id: string) => {
    if (!id) {
        throw '[insertSpanAtCursorPosition]: id must be supplied'
    }

    const range = getRangeAtCursorPosition()
    if (!range) {
        return null
    }

    const elem = document.createElement('span')
    elem.id = id
    range.insertNode(elem)

    return elem
}

export const insertTextAtCursorPosition = (text: string) => {
    if (!text) {
        return null
    }

    const range = getRangeAtCursorPosition()
    if (!range) {
        return null
    }

    const textNode = document.createTextNode(text)
    range.insertNode(textNode)
    range.setStartAfter(textNode)

    return textNode
}

export const removeElement = (element: string | HTMLElement): HTMLElement|null => {
    const elementToRemove = typeof element === 'string' ? document.getElementById(element) : element
    if (elementToRemove?.parentNode != null) {
        return elementToRemove.parentNode.removeChild(elementToRemove)
    }
    else {return null}
}

export const saveCursorPostion=(container: any) =>{
    var containerSel = container.ownerDocument;
    var cursorPosition = { start: 0, end: 0 };

    if (containerSel.getSelection && containerSel.createRange) {
        var range = containerSel.getSelection().getRangeAt(0);
            var preSelectionRange = range.cloneRange();
            preSelectionRange.selectNodeContents(container);
            preSelectionRange.setEnd(range.startContainer, range.startOffset);
            cursorPosition.start = preSelectionRange.toString().length;
            cursorPosition.end = cursorPosition.start + range.toString().length;
        
    } else if (containerSel.selection && containerSel.body.createTextRange) {
        var selectedTextRange = containerSel.selection.createRange();
        var preSelectionTextRange = containerSel.body.createTextRange();
        preSelectionTextRange.moveToElementText(container);
        preSelectionTextRange.setEndPoint("EndToStart", selectedTextRange);
        cursorPosition.start = preSelectionTextRange.text.length;
        cursorPosition.end = cursorPosition.start + selectedTextRange.text.length;
    }
    return cursorPosition;
};

export const cursorPositionRestore=(container: any, savedSel: any)=> {
    if (savedSel === null || savedSel === undefined) return false;

    var documentContainer = container.ownerDocument;
    var window = 'defaultView' in documentContainer ? documentContainer.defaultView : documentContainer.parentWindow;

    if (window.getSelection && documentContainer.createRange) {
        var charIndex = 0, range = documentContainer.createRange();
        range.setStart(container, 0);
        range.collapse(true);
        var nodeStack = [container], node, foundStart = false, stop = false;

        while (!stop && (node = nodeStack.pop())) {
            if (node.nodeType == 3) {
                var nextCharIndex = charIndex + node.length;
                if (!foundStart && savedSel.start >= charIndex && savedSel.start <= nextCharIndex) {
                    range.setStart(node, savedSel.start - charIndex);
                    foundStart = true;
                }
                if (foundStart && savedSel.end >= charIndex && savedSel.end <= nextCharIndex) {
                    range.setEnd(node, savedSel.end - charIndex);
                    stop = true;
                }
                charIndex = nextCharIndex;
            } else {
                var i = node.childNodes.length;
                while (i--) {
                    nodeStack.push(node.childNodes[i]);
                }
            }
        }
        var sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
    } else if (documentContainer.selection && documentContainer.body.createTextRange) {
        var textRange = documentContainer.body.createTextRange();
        textRange.moveToElementText(container);
        textRange.collapse(true);
        textRange.moveEnd("character", savedSel.end);
        textRange.moveStart("character", savedSel.start);
        textRange.select();
    }
};