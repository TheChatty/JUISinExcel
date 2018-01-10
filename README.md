This Excel macro contains VBA code to query AVM's JSON Update Info Service.

The left button on the only sheet will query all FRITZ!Boxes listed in the IP table (thus skip empty entries) and write their details in the box info table. You may have manual entries in the box info table if you don't have an IP entry next to it (like the third entry in the example).

The right button will query AVM's JUIS listed in the box info table (thus skip emptyx entries) and write the result (n/a or URL to the latest firmware image) in the update table. You may have manual entries in the update table if you don't have "juis?" entry next to it (makes no sense actually).

All tables can be extended line-wise. They should be of the same line count.