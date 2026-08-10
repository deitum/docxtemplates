type LogSink = (message?: string, ...optionalParams: unknown[]) => void;

const noSink: LogSink = () => {};

/**
 * Debug logging is off unless the host installs a sink. The `enabled` flag lets
 * hot paths skip building a message they would only throw away — serializing a
 * node costs real time once per node of a large document.
 */
export const logger: { debug: LogSink; enabled: boolean } = {
  debug: noSink,
  enabled: false,
};

/** Installs the sink debug messages go to, or removes it again with `null`. */
export function setDebugLogSink(sink: LogSink | null) {
  logger.debug = sink ?? noSink;
  logger.enabled = sink != null;
}
