set wmp=createobject("wmplayer.ocx")

set cd=wmp.cdromcollection.item(0)

do

cd.eject

loop