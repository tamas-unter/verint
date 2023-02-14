function Compare-XmlNodesPlain{
# todo: haschildnodes -- recurse
# todo: that has no such tag / element
# todo: the other way around
# todo: check attributes as well!
    param ([System.Xml.XmlElement]$this, [System.Xml.XmlElement]$that)
    $this.SelectNodes("*")|
        select -ExpandProperty Name|%{
            $v1=$this.GetElementsByTagName("$_").'#text'
            $v2=$that.GetElementsByTagName("$_").'#text'
            [pscustomobject]@{
                prop=$_
                this=$v1
                that=$v2
                result=($v1 -eq $v2)
            }
        }
}