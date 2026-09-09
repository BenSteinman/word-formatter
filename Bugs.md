1. ~~Headers not classified as bold in runs.~~ RESOLVED
2. Deep copy of tables drops manually applied cell shading but preserves table styles.
    - Tables with custom/named styles are preserved correctly
    - Manually shaded cells (direct formatting) lose their shading after deep copy
    - Needs investigation into how cell shading is stored in XML
3. Page numbers and footer text alignment linked.



<!-- TODO -->
- ~~Implement All caps for TO, FROM, DATE, RE~~ RESOLVED
- ~~Auto space correctly (tab alignment for memo header)~~ RESOLVED
- ~~Add top, side, bottom margin option~~ RESOLVED