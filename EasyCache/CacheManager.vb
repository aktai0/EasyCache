Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

<Serializable()>
Module CacheManager
   Private CacheTypeNames As New Dictionary(Of String, EasyCache)

   Private _CacheList As New Dictionary(Of String, EasyCache)
   Public Property CacheList() As Dictionary(Of String, EasyCache)
      Get
         Return _CacheList
      End Get
      Set(ByVal value As Dictionary(Of String, EasyCache))
         _CacheList = value
      End Set
   End Property

   Public Function RetrieveCache(Of T As EasyCache)() As T
      For Each cache As EasyCache In CacheList.Values
         If TypeOf (cache) Is T Then
            Return CType(cache, T)
         End If
      Next
      Return Nothing
   End Function

   ' We use two dicts because A) Dictionary entries are added as (Address, Address) pairs a type
   '  is not a primitive type and B) Deserialization creates a new object, so we would either need
   '  to replace the entry in the original dictionary (if we're only using one), use a different
   '  dictionary, or use a container class.
   Public Sub Init(ByVal caches As IEnumerable(Of EasyCache))
      For Each cache In caches
         If CacheTypeNames.ContainsKey(cache.CACHE_NAME) Then
            Throw New EasyCacheException("Error: cannot have more than one of the same cache (did you call EasyCache.Init() with a list that has two of the same cache names?")
         Else
            CacheTypeNames.Add(cache.CACHE_NAME, cache)
         End If
      Next
   End Sub

   Public Sub LoadAllCaches()
      For Each kvpair As KeyValuePair(Of String, EasyCache) In CacheTypeNames
         Dim loadedCache As EasyCache = LoadCacheFile(kvpair.Value.CACHE_FILE_NAME)
         ' If the cache file doesn't exist...
         If loadedCache Is Nothing Then
            loadedCache = kvpair.Value
         End If
         CacheList.Add(kvpair.Key, loadedCache)
         If loadedCache.ShouldRebuildCache Then
            loadedCache.RebuildCacheAsync()
         End If
         loadedCache.FinishedLoading()
      Next
   End Sub

   Private Function LoadCacheFile(ByVal cacheFileName As String, Optional ByVal messageBoxErrors As Boolean = True) As EasyCache
      Dim newCache As EasyCache = Nothing
      If File.Exists(cacheFileName) Then
         Dim TestFileStream As Stream = File.OpenRead(cacheFileName)
         Dim deserializer As New BinaryFormatter
         Try
            newCache = CType(deserializer.Deserialize(TestFileStream), EasyCache)
         Catch ex As Exception
            If messageBoxErrors Then
               MsgBox("There was an error with the cache. " & cacheFileName & " will be deleted. All cache data in that file will be lost. If you want to save a backup of it, do so NOW (before you exit this dialog). Error: " & ex.Message)
               TestFileStream.Close()
               File.Delete(cacheFileName)
            Else
               TestFileStream.Close()
               Throw New EasyCacheException("Error loading cache file: " & cacheFileName & ". If you're changing the class a lot, try deleting the cache file and trying again.")
            End If
         End Try
         TestFileStream.Close()
      End If
      Return newCache
   End Function

   Public Sub StoreAllCaches()
      For Each cache In CacheList.Values
         If cache.CacheChanged Then
            cache.StartedStoring()
            StoreCacheFile(cache.CACHE_FILE_NAME, cache)
         End If
      Next
   End Sub

   Public Sub StoreCacheFile(ByVal cacheFileName As String, ByRef givenCache As EasyCache)
      ' SyncLock to avoid corruption.
      SyncLock givenCache
         Dim TestFileStream As Stream = File.Create(cacheFileName)
         Dim serializer As New BinaryFormatter
         serializer.Serialize(TestFileStream, givenCache)
         TestFileStream.Close()
      End SyncLock
   End Sub
End Module

<Serializable()>
MustInherit Class EasyCache
   ' Name of the cache file (eg. "ImageCache.bin").
   MustOverride ReadOnly Property CACHE_FILE_NAME() As String

   ' Name of the cache key in Dictionary (eg. "Image").
   MustOverride ReadOnly Property CACHE_NAME() As String

   ' Always reuse old cache by default.
   Overridable ReadOnly Property ShouldRebuildCache() As Boolean
      Get
         Return False
      End Get
   End Property

   ' Always save by default.
   Overridable ReadOnly Property CacheChanged() As Boolean
      Get
         Return True
      End Get
   End Property

   ' Manual method for rebuilding. Not called by CacheManager but you might choose
   '  to implement this and call it in your own thread.
   Overridable Sub RebuildCache()
      Throw New InvalidOperationException
   End Sub

   ' Automatic method for rebuilding. Called by CacheManager if ShouldRebuildCache
   '  is true.
   Overridable Sub RebuildCacheAsync()
   End Sub

   Sub New()
   End Sub

   ' Called right after CacheManager finishes loading this cache.
   Overridable Sub FinishedLoading()
   End Sub

   ' Called right before CacheManager begins serializing this cache for storing.
   '  Update variables that indicate whether the cache was changed or not
   '  For example: if you had a "last item changed" variable, now would be
   '  the time to update it.
   Overridable Sub StartedStoring()
   End Sub

End Class