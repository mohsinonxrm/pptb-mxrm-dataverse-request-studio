import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';

/**
 * Shown at the top of each clause pane (Select / Filter / Orderby / Expand /
 * Count) when `$apply` is enabled on the parent request. Per the Dataverse
 * aggregate-data docs:
 *
 *   "The $select, $orderby, $expand, and $count options are not supported with
 *    the $apply option."
 *
 * `$filter` at the top level is also discouraged when `$apply` is on — filter
 * stages live inside the $apply pipeline. Only `$top` survives end-to-end.
 *
 * The banner is opt-in per clause so the user still sees their existing values
 * (the editor isn't actually disabled — switching $apply off would restore the
 * behaviour). We just want to make the trade-off obvious.
 */
export function ApplyOverridesBanner({ clause }: { clause: '$select' | '$filter' | '$orderby' | '$expand' | '$count' }) {
  // The $orderby case is special — Dataverse doesn't just ignore it, it
  // rejects $orderby on the aggregate alias columns:
  //   "The query node SingleValueOpenPropertyAccess is not supported."
  // (per /webapi/query/aggregate-data). Surface that explicitly.
  const isOrderby = clause === '$orderby';
  return (
    <MessageBar layout="multiline" intent={isOrderby ? 'error' : 'warning'} style={{ marginBottom: 14 }}>
      <MessageBarBody>
        <MessageBarTitle>
          {isOrderby
            ? '$orderby on aggregate output is not supported by Dataverse.'
            : `${clause} is ignored while $apply is active.`}
        </MessageBarTitle>
        {isOrderby ? (
          <>Dataverse rejects <code>$orderby</code> on the alias columns produced by <code>$apply</code> with
            <em> "The query node SingleValueOpenPropertyAccess is not supported."</em> Remove your <code>$orderby</code>
            entries, or disable <code>$apply</code> if you need server-side ordering.</>
        ) : (
          <><code>$apply</code> takes over the response shape — only <code>$top</code> remains alongside it.
            Output columns come from your groupby + aggregate aliases; row-level filtering moves into a
            <code>filter(…)/groupby(…)</code> stage. Disable <code>$apply</code> to use {clause} again.</>
        )}
      </MessageBarBody>
    </MessageBar>
  );
}
