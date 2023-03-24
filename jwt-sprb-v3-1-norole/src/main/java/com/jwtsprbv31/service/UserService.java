package com.jwtsprbv31.service;

import com.jwtsprbv31.dto.UserDTO;
import com.jwtsprbv31.entity.UserEntity;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;

public interface UserService {
    UserEntity get(Long id);
    UserEntity getByEmail(String email);
    Page<UserEntity> list(Pageable pageable);
    UserEntity create(UserDTO dto);
    UserEntity update(UserDTO dto);
    void delete(Long id);
}
