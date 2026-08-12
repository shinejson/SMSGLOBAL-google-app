// --- CACHING LAYER FOR PERFORMANCE OPTIMIZATION ---

/**
 * Get cached data or fetch from source
 * Implements caching strategy to reduce sheet reads
 * @param {string} cacheKey - Unique cache key
 * @param {function} fetchFunction - Function to fetch data if cache miss
 * @param {number} ttlSeconds - Time to live in seconds (default: 600 = 10 minutes)
 * @returns {any} Cached or freshly fetched data
 */
function getCachedData(cacheKey, fetchFunction, ttlSeconds) {
  ttlSeconds = ttlSeconds || 600; // Default 10 minutes
  
  try {
    const cache = CacheService.getScriptCache();
    const cachedValue = cache.get(cacheKey);
    
    if (cachedValue) {
      Logger.log('Cache HIT: ' + cacheKey);
      try {
        return JSON.parse(cachedValue);
      } catch (e) {
        Logger.log('Cache parse error: ' + e.message);
        // If parse fails, invalidate cache and fetch fresh
        cache.remove(cacheKey);
      }
    }
    
    // Cache miss - fetch fresh data
    Logger.log('Cache MISS: ' + cacheKey);
    const data = fetchFunction();
    
    // Store in cache
    try {
      const jsonData = JSON.stringify(data);
      // CacheService has 100KB limit per entry, check size
      if (jsonData.length < 100000) {
        cache.put(cacheKey, jsonData, ttlSeconds);
        Logger.log('Cached: ' + cacheKey + ' (expires in ' + ttlSeconds + 's)');
      } else {
        Logger.log('WARNING: Data too large to cache (' + jsonData.length + ' bytes): ' + cacheKey);
      }
    } catch (e) {
      Logger.log('Cache storage error: ' + e.message);
    }
    
    return data;
    
  } catch (e) {
    Logger.log('Cache system error: ' + e.message);
    // If caching fails, just fetch data without caching
    return fetchFunction();
  }
}

/**
 * Invalidate (remove) cached data
 * Call this when data is modified (add, update, delete)
 * @param {string} cacheKey - Cache key to invalidate
 */
function invalidateCache(cacheKey) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(cacheKey);
    Logger.log('Cache invalidated: ' + cacheKey);
  } catch (e) {
    Logger.log('Cache invalidation error: ' + e.message);
  }
}

/**
 * Invalidate multiple cache keys at once
 * @param {array} cacheKeys - Array of cache keys to invalidate
 */
function invalidateMultipleCache(cacheKeys) {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(cacheKeys);
    Logger.log('Cache invalidated: ' + cacheKeys.join(', '));
  } catch (e) {
    Logger.log('Batch cache invalidation error: ' + e.message);
  }
}

/**
 * Clear all cached data
 * Use with caution - forces fresh fetch for all subsequent requests
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    // Get all known cache keys
    const knownKeys = [
      'data_Students',
      'data_Users',
      'data_Courses',
      'data_Classes',
      'data_Teachers',
      'data_Attendance',
      'data_Performance',
      'data_Invoices',
      'data_Payments',
      'data_Billings',
      'data_BillingCategories',
      'data_Parents',
      'data_Enrollments',
      'data_AcademicYears',
      'data_Permissions',
      'dashboard_stats'
    ];
    
    cache.removeAll(knownKeys);
    Logger.log('All cache cleared: ' + knownKeys.length + ' keys');
    
    return { success: true, cleared: knownKeys.length };
  } catch (e) {
    Logger.log('Clear all cache error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Get cache statistics (for monitoring)
 * @returns {object} Cache performance stats
 */
function getCacheStats() {
  // Note: CacheService doesn't provide built-in stats
  // This is a placeholder for custom implementation
  return {
    note: 'CacheService does not provide hit/miss statistics',
    recommendation: 'Monitor via Logger.log() entries',
    ttl: '600 seconds (10 minutes)',
    maxSize: '100KB per entry',
    totalCache: '10MB total'
  };
}

/**
 * Cached version of getStudentsData
 * @returns {array} Student data (cached)
 */
function getCachedStudentsData() {
  return getCachedData('data_Students', function() {
    // Call original function
    return getStudentsDataUncached();
  }, 600); // 10 minutes
}

/**
 * Cached version of getUsersData
 * @returns {array} User data (cached)
 */
function getCachedUsersData() {
  return getCachedData('data_Users', function() {
    return getUsersDataUncached();
  }, 600);
}

/**
 * Cached version of getCoursesData
 * @returns {array} Course data (cached)
 */
function getCachedCoursesData() {
  return getCachedData('data_Courses', function() {
    return getCoursesDataUncached();
  }, 600);
}

/**
 * Cached version of getClassesData
 * @returns {array} Class data (cached)
 */
function getCachedClassesData() {
  return getCachedData('data_Classes', function() {
    return getClassesDataUncached();
  }, 600);
}

/**
 * Cached version of dashboard stats
 * @returns {object} Dashboard statistics (cached)
 */
function getCachedDashboardStats() {
  return getCachedData('dashboard_stats', function() {
    return getDashboardStats();
  }, 300); // 5 minutes for dashboard (more frequent updates)
}

/**
 * Helper to wrap any data fetch function with caching
 * @param {string} dataType - Type of data (Students, Users, etc.)
 * @param {function} fetchFunc - Original fetch function
 * @param {number} ttl - Cache TTL in seconds
 * @returns {any} Cached or fresh data
 */
function getCachedDataGeneric(dataType, fetchFunc, ttl) {
  const cacheKey = 'data_' + dataType;
  return getCachedData(cacheKey, fetchFunc, ttl || 600);
}

/**
 * Smart cache warming - preload frequently used data
 * Call this during off-peak hours or on app startup
 */
function warmCache() {
  Logger.log('Starting cache warming...');
  
  const startTime = new Date().getTime();
  
  try {
    // Preload most frequently accessed data
    getCachedStudentsData();
    getCachedUsersData();
    getCachedCoursesData();
    getCachedClassesData();
    getCachedDashboardStats();
    
    const elapsed = new Date().getTime() - startTime;
    Logger.log('Cache warming completed in ' + elapsed + 'ms');
    
    return {
      success: true,
      elapsed: elapsed,
      message: 'Cache warmed successfully'
    };
  } catch (e) {
    Logger.log('Cache warming error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Automatic cache invalidation on data modification
 * Call this helper after any add/update/delete operation
 * @param {string} dataType - Type of data modified (Students, Users, etc.)
 */
function invalidateCacheOnModify(dataType) {
  const keysToInvalidate = ['data_' + dataType];
  
  // Also invalidate dashboard if data affects it
  const dashboardRelatedTypes = ['Students', 'Courses', 'Attendance', 'Invoices', 'Payments', 'Performance'];
  if (dashboardRelatedTypes.indexOf(dataType) !== -1) {
    keysToInvalidate.push('dashboard_stats');
  }
  
  invalidateMultipleCache(keysToInvalidate);
}

/**
 * Cache-aware wrapper for data operations
 * Automatically handles cache invalidation
 */
var CachedDataOperations = {
  /**
   * Add student with automatic cache invalidation
   */
  addStudent: function(studentData) {
    const result = addStudent(studentData);
    if (result.success) {
      invalidateCacheOnModify('Students');
    }
    return result;
  },
  
  /**
   * Update student with automatic cache invalidation
   */
  updateStudent: function(studentId, studentData) {
    const result = updateStudent(studentId, studentData);
    if (result.success) {
      invalidateCacheOnModify('Students');
    }
    return result;
  },
  
  /**
   * Delete student with automatic cache invalidation
   */
  deleteStudent: function(studentId) {
    const result = deleteStudent(studentId);
    if (result.success) {
      invalidateCacheOnModify('Students');
    }
    return result;
  },
  
  /**
   * Add user with automatic cache invalidation
   */
  addUser: function(userData) {
    const result = addUser(userData);
    if (result.success) {
      invalidateCacheOnModify('Users');
    }
    return result;
  },
  
  /**
   * Update user with automatic cache invalidation
   */
  updateUser: function(userId, userData) {
    const result = updateUser(userId, userData);
    if (result.success) {
      invalidateCacheOnModify('Users');
    }
    return result;
  },
  
  /**
   * Delete user with automatic cache invalidation
   */
  deleteUser: function(userId) {
    const result = deleteUser(userId);
    if (result.success) {
      invalidateCacheOnModify('Users');
    }
    return result;
  }
};

/**
 * Performance monitoring - measure cache effectiveness
 * @param {string} operation - Operation name
 * @param {function} func - Function to measure
 * @returns {object} Result with timing information
 */
function measurePerformance(operation, func) {
  const startTime = new Date().getTime();
  
  try {
    const result = func();
    const elapsed = new Date().getTime() - startTime;
    
    Logger.log('Performance: ' + operation + ' completed in ' + elapsed + 'ms');
    
    return {
      success: true,
      data: result,
      elapsed: elapsed
    };
  } catch (e) {
    const elapsed = new Date().getTime() - startTime;
    Logger.log('Performance: ' + operation + ' failed after ' + elapsed + 'ms - ' + e.message);
    
    return {
      success: false,
      error: e.message,
      elapsed: elapsed
    };
  }
}

/**
 * Compare cached vs uncached performance
 * @returns {object} Performance comparison results
 */
function benchmarkCache() {
  Logger.log('Starting cache benchmark...');
  
  // Clear cache first
  clearAllCache();
  
  // Measure uncached fetch
  const uncachedTime = measurePerformance('Uncached Students Fetch', function() {
    return getStudentsDataUncached();
  });
  
  // Measure cached fetch (first time - cache miss)
  const firstCachedTime = measurePerformance('First Cached Fetch (cache miss)', function() {
    return getCachedStudentsData();
  });
  
  // Measure cached fetch (second time - cache hit)
  const secondCachedTime = measurePerformance('Second Cached Fetch (cache hit)', function() {
    return getCachedStudentsData();
  });
  
  const improvement = Math.round((uncachedTime.elapsed / secondCachedTime.elapsed) * 100) / 100;
  
  return {
    uncached: uncachedTime.elapsed + 'ms',
    cachedMiss: firstCachedTime.elapsed + 'ms',
    cachedHit: secondCachedTime.elapsed + 'ms',
    speedup: improvement + 'x faster',
    recommendation: improvement > 2 ? 'Caching highly effective' : 'Consider longer TTL or data structure optimization'
  };
}
