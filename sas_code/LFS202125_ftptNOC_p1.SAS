DATA work.lfstats5digNOC;
SET RTRAdata.LFS202125 (keep = ID PROV AGE LFSSTAT FTPTMAIN NOC_5 SYEAR);

/* Province B.C.*/
if PROV = 59;

/* Age 15 or over */
if AGE >= 15;

/*Labour Force Status*/

If LFSSTAT IN (1,2);

/*First half of the NOCS plus missing*/

if NOC_5 < 50000;
if missing(NOC_5) then NOC_5="missing";

/* Part-time and Full-time status */
if ftptmain = '1' then ftpt =      'Full-time';
else if ftptmain = '2' then ftpt = 'Part-time';
else ftpt = 'Unknown';

run;

%RTRAFreq(
     InputDataset=work.lfstats5digNOC,
     OutputName=ftpt2125NOCp1,
     ClassVarList= SYEAR ftpt NOC_5,
     UserWeight=FINALWT);
run;
