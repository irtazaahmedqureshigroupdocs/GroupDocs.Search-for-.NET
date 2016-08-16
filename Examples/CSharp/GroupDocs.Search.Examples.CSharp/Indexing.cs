﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GroupDocs.Search;
using GroupDocs.Search.Events;


namespace GroupDocs.Search_for_.NET
{
    class Indexing
    {
        /// <summary>
        /// Update index
        /// </summary>
        public static void UpdateIndex()
        {
            //ExStart:UpdateIndex
            // Load index
            Index index = new Index(Utilities.indexPath);
            index.Update();
            //ExEnd:UpdateIndex
        }

        /// <summary>
        /// Update index repository
        /// </summary>
        public static void UpdateIndexInRepository()
        {
            //ExStart:UpdateIndexInRepository
            IndexRepository repository = new IndexRepository();
            repository.AddToRepository(Utilities.indexPath);
            repository.AddToRepository(Utilities.indexPath2);
            // Update all indexes in repository
            repository.Update();
            //ExEnd:UpdateIndexInRepository
        }

        /// <summary>
        /// Update index asynchronously
        /// </summary>
        public static void UpdateIndexAsynchronously()
        {
            //ExStart:UpdateIndexAsynchronously
            //Load index
            Index index = new Index(Utilities.indexPath);
            index.OperationFinished += Utilities.index_OperationFinished;
            // Update index asynchronously
            index.UpdateAsync();
            //ExEnd:UpdateIndexAsynchronously
        }

        /// <summary>
        /// Update index in repository asynchronously
        /// </summary>
        public static void UpdateIndexInRepoAsynchronously()
        {
            //ExStart:UpdateIndexInRepoAsynchronously
            IndexRepository repository = new IndexRepository();
            repository.OperationFinished += Utilities.index_OperationFinished;

            repository.AddToRepository(Utilities.indexPath);
            repository.AddToRepository(Utilities.indexPath2);

            // Update all indexes in repository asynchronously
            repository.UpdateAsync();
            //ExEnd:UpdateIndexInRepoAsynchronously
        }

        /// <summary>
        /// Create index in memory
        /// </summary>
        public static void CreateIndexInMemory()
        {
            //ExStart:CreateIndexInMemory
            // Create index in memory
            Index index1 = new Index();
            // Create index in memory using Index Repository
            IndexRepository repository = new IndexRepository();
            Index index2 = repository.Create();
            //ExEnd:CreateIndexInMemory
        }

        /// <summary>
        /// Create index on disk
        /// </summary>
        public static void CreateIndexOnDisk()
        {
            //ExStart:CreateIndexOnDisk
            // Create index on disk
            Index index1 = new Index(Utilities.indexPath);
            // Create index on disk using Index Repository
            IndexRepository repository = new IndexRepository();
            Index index2 = repository.Create(Utilities.indexPath2);
            //ExEnd:CreateIndexOnDisk
        }

        /// <summary>
        /// Create index in memory with index settings
        /// </summary>
        public static void CreateIndexInMemoryWithIndexSettings()
        {
            //ExStart:CreateIndexInMemoryWithIndexSettings
            bool quickIndexing = true;
            IndexingSettings settings = new IndexingSettings(quickIndexing);

            // Create index on disk
            Index index1 = new Index(settings);

            // Create index on disk using Index Repository
            IndexRepository repository = new IndexRepository();
            Index index2 = repository.Create(settings);
            //ExEnd:CreateIndexInMemoryWithIndexSettings
        }

        /// <summary>
        /// Create with overwriting existed index
        /// </summary>
        public static void CreateWithOverwritingExistedIndex()
        {
            //ExStart:CreateWithOverwritingExistedIndex
            // Create index on disk. If Index folder is not empty it will be rewited
            Index index1 = new Index(Utilities.indexPath, true);

            // Create index on disk using Index Repository
            IndexRepository repository = new IndexRepository();
            Index index2 = repository.Create(Utilities.indexPath);
            //ExEnd:CreateWithOverwritingExistedIndex
        }

        /// <summary>
        /// Load index
        /// </summary>
        public static void LoadIndex()
        {
            //ExStart:loadindex
            // Load index
            Index index = new Index(Utilities.indexPath);

            // Load indexes to index repository
            IndexRepository repository = new IndexRepository();
            repository.AddToRepository(index);
            //repository.AddToRepository(Utilities.indexPath2);
            //ExEnd:loadindex
        }

        /// <summary>
        /// Add document to index
        /// </summary>
        public static void AddDocumentToIndex()
        {
            //ExStart:AddDocumentToIndex
            // Create index
            Index index = new Index(Utilities.indexPath);
            // all files from folder and its subfolders will be added to the index
            index.AddToIndex(Utilities.documentsPath);
            //ExEnd:AddDocumentToIndex
        }

        /// <summary>
        /// Add document to index asynchronously
        /// </summary>
        public static void AddDocumentToIndexAsynchronously()
        {
            //ExStart:AddDocumentToIndexAsynchronously
            // Create index
            Index index = new Index(Utilities.indexPath);
            index.OperationFinished += Utilities.index_OperationFinished;
            // all files from folder and its subfolders will be added to the index
            index.AddToIndexAsync(Utilities.documentsPath);
            //ExEnd:AddDocumentToIndexAsynchronously
        }

        /// <summary>
        /// Subscription to events
        /// </summary>
        public static void SubscriptionToEvents()
        {
            //ExStart:SubscriptionToEvents
            // Create index in memory
            Index index = new Index();
            index.OperationFinished += Utilities.index_OperationFinished;
            index.AddToIndexAsync(Utilities.documentsPath);
            index.UpdateAsync();
            //ExEnd:SubscriptionToEvents
        }

        /// <summary>
        /// Custom extractor test
        /// </summary>
        public static void CustomExtractor()
        {
            //ExStart:CustomExtractor
            Index index = new Index(Utilities.indexPath);
            index.CustomExtractors.Add(new CustomFieldExtractor());

            index.AddToIndex(Utilities.documentsPath);
            //ExEnd:CustomExtractor
        }


        /// <summary>
        /// Add PowerPoint Document to index
        /// </summary>
        public static void AddPowerPointDocumentToIndex()
        {
            //ExStart:AddPowerPointDocumentToIndex
            // Create index
            Index index = new Index(Utilities.indexPath);
            // all files from folder and its subfolders will be added to the index
            index.AddToIndex(Utilities.documentsPath);

            SearchResults results1 = index.Search("author:cisco"); // searching by author of presentation
            SearchResults results2 = index.Search("LastSavedBy:teresa"); // searching by person who saved presentation last time
            //ExEnd:AddPowerPointDocumentToIndex
        }

        /// <summary>
        /// Prevents Unnecessary File Indexing
        /// </summary>
        public static void PreventUnnecessaryFileIndex()
        {
            //ExStart: PreventUnnecessaryFileIndex
            // Create index
            Index index = new Index(Utilities.indexPath);

            // Add documents to index
            index.AddToIndex(Utilities.documentsPath);

            // Try add the same documents to index
            index.AddToIndex(Utilities.documentsPath); // Already indexed files will not be reindexed.
            //ExEnd: PreventUnnecessaryFileIndex
        }

        /// <summary>
        /// Search and Browse Email using Aspose.Email API
        /// </summary>
        public static void SearchingEmailMessages()
        {
            //ExStart: SearchingEmailMessages
            // Create index
            Index index = new Index(Utilities.indexPath);

            // Indexing MS Outlook storage with email messages
            index.OperationFinished += index_OperationFinished;
            index.ErrorHappened += index_ErrorHappened;
            index.AddToIndex(Utilities.documentsPath);

            // Searching in index
            SearchResults results = index.Search(searchQuery);

            // User gets all messages that qualify to search query using Aspose.Email API
            MessageInfoCollection messages = new MessageInfoCollection();
            foreach (DocumentResultInfo searchResult in results)
            {
                if (searchResult.DocumentType == DocumentType.OutlookEmailMessage)
                {
                    OutlookEmailMessageResultInfo emailResultInfo = searchResult as OutlookEmailMessageResultInfo;
                    MessageInfo message = GetEmailMessagesById(pstFileLink, emailResultInfo.EntryIdString);
                    if (message != null)
                    {
                        messages.Add(message);
                    }
                }
            }
            //ExEnd: SearchingEmailMessages
        }

        #region Getting Email Messages by EntryIdString using Aspose.Email API

        private MessageInfo GetEmailMessagesById(string fileName, string fieldId)
        {
            PersonalStorage pst = PersonalStorage.FromFile(fileName, false);
            return GetEmailMessagesById(pst.RootFolder, fieldId);
        }

        private MessageInfo GetEmailMessagesById(FolderInfo folderInfo, string fieldId)
        {
            MessageInfo result = null;
            MessageInfoCollection messageInfoCollection = folderInfo.GetContents();
            foreach (MessageInfo messageInfo in messageInfoCollection)
            {
                if (messageInfo.EntryIdString == fieldId)
                {
                    result = messageInfo;
                    break;
                }
            }

            if (result == null && folderInfo.HasSubFolders)
            {
                foreach (FolderInfo subfolderInfo in folderInfo.GetSubFolders())
                {
                    result = GetEmailMessagesById(subfolderInfo, fieldId);
                    if (result != null)
                    {
                        break;
                    }
                }
            }
            return result;
        }

        #endregion
    }
}
