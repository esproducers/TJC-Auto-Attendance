def update_cache(self, new_image_path):
    """
    增量更新缓存：只处理新加入的一张图片
    """
    if not hasattr(self, 'cache') or self.cache is None:
        # 若缓存未初始化，则全量构建（仅在程序启动时）
        self.build_cache()
        return
    
    # 读取新图片并提取特征（假设每张照片只有一个人脸）
    img = cv2.imread(new_image_path)
    if img is None:
        raise ValueError(f"Cannot read image: {new_image_path}")
    
    faces = self.model.get(img)  # 使用 InsightFace 提取
    if len(faces) == 0:
        raise ValueError("No face detected in the new image")
    
    # 取第一个人脸
    embedding = faces[0].normed_embedding
    # 根据文件名生成 key
    name = os.path.splitext(os.path.basename(new_image_path))[0]
    # 存入缓存字典
    self.cache[name] = embedding
    
    # 保存缓存到文件
    with open(self.cache_path, 'wb') as f:
        pickle.dump(self.cache, f)
    print(f"[CACHE] Updated cache with {name}")