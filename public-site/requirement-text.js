// Excel alignment padding is not Word layout. Keep units and line boundaries.
export const requirementText=value=>String(value??'').split(/\r?\n/).map(line=>line.replace(/[\t \u00a0\u3000]+/g,' ').trim()).join('\n');
export const manufacturingFields=new Set(['production','material','layers','thickness','copper','baseCopper','mask','legend','finish','inspection','size','panel','stackup','via','impedance','ul','origin','reports','goldfinger']);
