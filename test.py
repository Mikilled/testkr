from lxml import etree

xml = '''
<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:nvPicPr>
                <pic:cNvPr name="" id="1"/>
                <pic:cNvPicPr/>
            </pic:nvPicPr>
            <pic:blipFill>
                <a:blip r:embed="rId20"/>
                <a:stretch>
                    <a:fillRect/>
                </a:stretch>
            </pic:blipFill>
            <pic:spPr>
                <a:xfrm>
                    <a:off y="0" x="0"/>
                    <a:ext cx="6624320" cy="4269740"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
            </pic:spPr>
        </pic:pic>
    </a:graphicData>
</a:graphic>
'''

root = etree.fromstring(xml)
blip = root.xpath("//a:blip", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})[0]
rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
ext = root.xpath(f"//a:ext[preceding-sibling::a:blip[@r:embed='{rId}']]", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})[0]
cx = ext.get("cx")
cy = ext.get("cy")

print("cx:", cx)
print("cy:", cy)
