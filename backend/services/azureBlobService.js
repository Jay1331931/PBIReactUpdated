const { BlobServiceClient, generateBlobSASQueryParameters, BlobSASPermissions, StorageSharedKeyCredential } = require("@azure/storage-blob");

const ACCOUNT_NAME   = process.env.AZURE_STORAGE_NAME;                  
const ACCOUNT_KEY    = process.env.AZURE_STORAGE_ACCESS_KEY;      
const CONTAINER_NAME = process.env.AZURE_STORAGE_EXPORT_CONTAINER;

const sharedKeyCredential = new StorageSharedKeyCredential(ACCOUNT_NAME, ACCOUNT_KEY);
const blobServiceClient   = new BlobServiceClient(
  `https://${ACCOUNT_NAME}.blob.core.windows.net`,
  sharedKeyCredential
);

async function uploadAndGetSasUrl(buffer, filename, mimeType, expiryMinutes = 30) {
  const containerClient = blobServiceClient.getContainerClient(CONTAINER_NAME);
  // await containerClient.createIfNotExists({ access: "private" });

  const blobName  = `exports/${Date.now()}-${filename}`;
  const blockBlob = containerClient.getBlockBlobClient(blobName);

  await blockBlob.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType:        mimeType,
      blobContentDisposition: `attachment; filename="${filename}"`,
    },
  });

  const expiresOn = new Date(Date.now() + expiryMinutes * 60 * 1000);
  const sasToken  = generateBlobSASQueryParameters(
    {
      containerName: CONTAINER_NAME,
      blobName,
      permissions:   BlobSASPermissions.parse("r"),
      expiresOn,
    },
    sharedKeyCredential
  ).toString();

  return {
    sasUrl:   `${blockBlob.url}?${sasToken}`,
    blobName,
  };
}

async function deleteBlob(blobName) {
  const containerClient = blobServiceClient.getContainerClient(CONTAINER_NAME);
  await containerClient.getBlockBlobClient(blobName).deleteIfExists();
}

module.exports = { uploadAndGetSasUrl, deleteBlob };