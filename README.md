# excel_to_neo
## Created on Mon Jan 9 10:59:37 2017
### author: ryancollingwood@gmail.com

Script for importing tabular excel document into a Neo4j graph database.
The intension is to import not only cell value but the additional information, such as:
* Formulas
* Cell Decorations - Colour, Borders
* Font Decorations - Colour, Family, Modifiers

**At present only the cell values are imported.**

Required Libraries
* Neo4j - https://neo4j.com
* Neo4j Python Driver Soruce - https://github.com/neo4j/neo4j-python-driver
* xLwings - https://www.xlwings.org/
* xlwings Source - https://github.com/ZoomerAnalytics/xlwings

Function `export_sheet` is the entry point function

Some quick improvements that could be made:

* Move the consts to an external configuration file;
* Find a better solution to reading through blank rows, rather using stopping after C_MAX_EMPTY blank rows;
* Use Pandas Dataframe for creating an in memory representation of the Spreadsheet;
* Use Transaction for handling upserts, so that rollback is possible;

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.