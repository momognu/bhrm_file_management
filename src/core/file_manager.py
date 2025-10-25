import os


class FileManager:
    def __init__(self):
        pass
        
    def load_files_tree(self, directory):
        """加载目录中的所有文件，组织成树状结构"""
        tree = {
            'name': os.path.basename(directory),
            'path': directory,
            'type': 'directory',
            'children': []
        }
        
        # 遍历目录及其子目录
        for root, dirs, filenames in os.walk(directory):
            # 创建目录结构
            for dirname in dirs:
                dir_path = os.path.join(root, dirname)
                relative_path = os.path.relpath(dir_path, directory)
                dir_node = {
                    'name': dirname,
                    'path': dir_path,
                    'type': 'directory',
                    'children': []
                }
                
                # 找到父节点并添加
                parent_node = self._find_parent_node(tree, relative_path)
                if parent_node:
                    parent_node['children'].append(dir_node)
            
            # 添加文件
            for filename in filenames:
                file_path = os.path.join(root, filename)
                relative_path = os.path.relpath(file_path, directory)
                try:
                    stat = os.stat(file_path)
                    file_node = {
                        'name': filename,
                        'size': stat.st_size,
                        'modified': stat.st_mtime,
                        'path': file_path,
                        'type': 'file',
                        'extension': os.path.splitext(filename)[1] if os.path.splitext(filename)[1] else '文件'
                    }
                    
                    # 找到父节点并添加
                    parent_node = self._find_parent_node(tree, relative_path)
                    if parent_node:
                        parent_node['children'].append(file_node)
                except Exception as e:
                    print(f"无法读取文件信息 {file_path}: {e}")
                    
        return tree
        
    def _find_parent_node(self, tree, relative_path):
        """找到相对路径对应的父节点"""
        if relative_path == '.':
            return tree
            
        parts = relative_path.split(os.sep)
        current_node = tree
        
        # 遍历路径部分，找到对应的节点
        for part in parts[:-1]:  # 最后一个是文件名或目录名，不包含在路径中
            found = False
            for child in current_node['children']:
                if child['name'] == part and child['type'] == 'directory':
                    current_node = child
                    found = True
                    break
            if not found:
                return None
                
        return current_node
        
    def format_size(self, size):
        """格式化文件大小"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} PB"