// load in u.q (available on kx wiki)
\d .u
init:{w::t!(count t::tables`.)#()}
del:{w[x]_:w[x;;0]?y};.z.pc:{del[;x]each t};
sel:{$[`~y;x;select from x where sym in y]}
pub:{[t;x]{[t;x;w]if[count x:sel[x]w 1;(neg first w)(`upd;t;x)]}[t;x]each w t}
add:{$[(count w x)>i:w[x;;0]?.z.w;.[`.u.w;(x;i;1);union;y];w[x],:enlist(.z.w;y)];(x;$[99=type v:value x;sel[v]y;0#v])}
sub:{if[x~`;:sub[;y]each t];if[not x in t;'x];add[x;y]}
end:{(neg union/[w[;;0]])@\:(`.u.end;x)}
\d .

// create the table to be published (subscription key; value)
// tables to be published require a sym column, which can be of any type
xlsub:([]sym:`$(); val:())

// initialise pubsub
// all tables in the top level namespace (`.) become publish-able
// tables that can be published can be seen in .u.w
// in our case excel clients will all use the same table, tsub
.u.init[]

// functions to publish data
// .u.pub takes the table name and table data
.xl.getdata:{([]sym:enlist `key;val:1?100f)}
.xl.publish:{.u.pub[`xlsub; .xl.getdata[]]}

// utils
.xl.dt:{`datetime$x-2000.01.01-1900.01.01-2} /convert an excel date into a q datetime

// create timer function to randomly publish
.z.ts:{.xl.publish[]}

// fire timer every N ms
\t 0