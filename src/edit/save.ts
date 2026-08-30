import type { App, TFile } from "obsidian";
import type { PptxPackage } from "../pptx/package";

export interface SaveResult {
	/** Path of the backup written by this save, if one was created. */
	backupPath: string | null;
	bytesWritten: number;
}

export class ConflictError extends Error {
	constructor() {
		super("The file changed on disk since it was opened. Reload it before saving.");
		this.name = "ConflictError";
	}
}

/**
 * Write an edited deck back to the vault.
 *
 * The first save of a file also leaves a one-off `.pptx.bak` beside it. Binary
 * files do not go through Obsidian's file recovery, so without that copy a bad
 * round trip would be unrecoverable.
 */
export async function saveDeck(
	app: App,
	file: TFile,
	pkg: PptxPackage,
	expectedMtime: number,
): Promise<SaveResult> {
	if (file.stat.mtime !== expectedMtime) throw new ConflictError();

	const backupPath = `${file.path}.bak`;
	let createdBackup: string | null = null;
	if (!app.vault.getAbstractFileByPath(backupPath)) {
		const original = await app.vault.readBinary(file);
		await app.vault.createBinary(backupPath, original);
		createdBackup = backupPath;
	}

	const bytes = pkg.toZip();
	const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
	await app.vault.modifyBinary(file, buffer as ArrayBuffer);
	pkg.clearDirty();

	return { backupPath: createdBackup, bytesWritten: bytes.byteLength };
}
