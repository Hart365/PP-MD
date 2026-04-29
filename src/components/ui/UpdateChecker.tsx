/**
 * @file UpdateChecker.tsx
 * @description Component for checking and displaying app updates.
 */

import { useState } from 'react';
import { useUpdateCheck } from '../../hooks/useUpdateCheck';
import { formatFileSize } from '../../utils/versionUtils';
import styles from './UpdateChecker.module.css';

/**
 * UpdateChecker component with button and modal for displaying updates
 */
export function UpdateChecker() {
  const [showDialog, setShowDialog] = useState(false);
  const { checking, result, error, checkForUpdates, dismiss } = useUpdateCheck();

  const handleCheck = async () => {
    setShowDialog(true);
    await checkForUpdates();
  };

  const handleDismiss = () => {
    dismiss();
    setShowDialog(false);
  };

  const handleDownload = () => {
    if (result?.downloadAsset?.downloadUrl) {
      window.open(result.downloadAsset.downloadUrl, '_blank');
    } else if (result?.releaseUrl) {
      window.open(result.releaseUrl, '_blank');
    }
  };

  return (
    <>
      {/* Update Check Button */}
      <button
        className={styles.updateButton}
        onClick={handleCheck}
        disabled={checking}
        title="Check for available updates"
        aria-label="Check for updates"
      >
        {checking ? '⟳ Checking...' : '⟳ Updates'}
      </button>

      {/* Update Dialog */}
      {showDialog && (
        <div className={styles.dialogOverlay} onClick={handleDismiss}>
          <dialog
            className={styles.dialog}
            onClick={e => e.stopPropagation()}
            open
          >
            <div className={styles.dialogContent}>
              <h2 className={styles.dialogTitle}>Check for Updates</h2>

              {checking && (
                <div className={styles.loading}>
                  <div className={styles.spinner}></div>
                  <p>Checking for updates...</p>
                </div>
              )}

              {error && (
                <div className={styles.error}>
                  <p className={styles.errorTitle}>Error checking for updates</p>
                  <p className={styles.errorMessage}>{error}</p>
                </div>
              )}

              {result && !checking && (
                <div className={styles.result}>
                  <div className={styles.versionInfo}>
                    <p>
                      <strong>Current version:</strong> {result.currentVersion}
                    </p>
                    <p>
                      <strong>Latest version:</strong> {result.latestVersion}
                    </p>
                    {result.releaseDate && (
                      <p>
                        <strong>Released:</strong>{' '}
                        {new Date(result.releaseDate).toLocaleDateString()}
                      </p>
                    )}
                  </div>

                  {result.hasUpdate ? (
                    <div className={styles.updateAvailable}>
                      <div className={styles.updateBanner}>
                        ✓ A new version is available!
                      </div>
                      {result.downloadAsset && (
                        <div className={styles.downloadInfo}>
                          <p>
                            <strong>File:</strong> {result.downloadAsset.name}
                          </p>
                          <p>
                            <strong>Size:</strong>{' '}
                            {formatFileSize(result.downloadAsset.size)}
                          </p>
                        </div>
                      )}
                    </div>
                  ) : (
                    <div className={styles.noUpdate}>
                      <p>✓ You are already using the latest version!</p>
                    </div>
                  )}
                </div>
              )}

              {/* Dialog Buttons */}
              <div className={styles.dialogButtons}>
                {result && result.hasUpdate && result.downloadAsset && (
                  <button
                    className={`${styles.button} ${styles.downloadButton}`}
                    onClick={handleDownload}
                  >
                    Download Update
                  </button>
                )}
                {result && !result.hasUpdate && (
                  <button
                    className={`${styles.button} ${styles.primaryButton}`}
                    onClick={handleDismiss}
                  >
                    OK
                  </button>
                )}
                {error && (
                  <button
                    className={`${styles.button} ${styles.primaryButton}`}
                    onClick={handleDismiss}
                  >
                    Close
                  </button>
                )}
                {!result && !error && (
                  <button
                    className={`${styles.button} ${styles.secondaryButton}`}
                    onClick={handleDismiss}
                    disabled={checking}
                  >
                    Cancel
                  </button>
                )}
              </div>
            </div>
          </dialog>
        </div>
      )}
    </>
  );
}
