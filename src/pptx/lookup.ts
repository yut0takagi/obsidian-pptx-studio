/**
 * A value from a lookup table, but only one the table itself defines.
 *
 * The keys these tables are read with come out of the deck — an element name,
 * an attribute value, a part's file extension — and a plain object answers
 * "constructor", "toString" or "__proto__" with what it inherits from
 * Object.prototype: a function or an object where a string or a number was
 * expected. Going through here, such a name is simply not in the table.
 */
export function own<T>(table: Readonly<Record<string, T>>, key: string | null | undefined): T | undefined {
	if (key === null || key === undefined) return undefined;
	return Object.prototype.hasOwnProperty.call(table, key) ? table[key] : undefined;
}
